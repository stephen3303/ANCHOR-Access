Option Compare Database
Option Explicit

' Access UI kit - the shared Personnel directory: who a person is, which apps
' they may open, and signing them in.
'
' The directory is one encrypted data file, Personnel_be.accdb, in its own
' restricted folder on the share, that every app links to. It holds these
' tables, and nothing else holds people's names:
'   kit_person      one row per person: Windows account and SID, sign-in
'                   name, rank and name, command, email, app password hash,
'                   directory admin, active (no SSN)
'   kit_person_app  which apps (app_code "SABER", "SPM", "MG") a person may open
'   kit_person_upn  a person's user principal name ("1234567890.A@smil.mil",
'                   so it holds their DoD ID), kept to find their Windows
'                   account again. Made the first time one is saved. It is
'                   never linked into a front end: only an active admin
'                   reads or writes it (PersonnelUpn, PersonnelSetUpn,
'                   Import), and it is never exported or written to a log.
' Each app keeps its own roles and links them to person_id.
'
' Startup, straight after the app's own KitEnsureBackend:
'   Select Case PersonnelStatus("item")          ' item: one of the app's linked tables
'       Case PERSONNEL_READY: signed = PersonnelSignInWindows(APP_CODE)
'       Case Else: ... the setup screen (find the file, create it, or ask the admin)
'   End Select
' PersonnelStatus links the file by itself when it is in the default place:
' a Personnel folder beside the app's data folder (S:\Apps\Data\app.accdb
' gives S:\Apps\Personnel\Personnel_be.accdb).
'
' Signing in sets TempVars PersonID, Username (what KitUser() stamps),
' PersonName ("LT J. Smith") and PersonIsAdmin. PersonnelSignInWindows
' matches the Windows account (WinSid, then WinAccount) with no questions;
' after Sign out, PersonnelSignIn checks a sign-in name and app password.
'
' The file's password is the site key, chosen by the first admin when they
' create the file and kept in Personnel.key beside it, never in the code.
' Every app reads the key from there when it finds or links the file
' (PersonnelKeyFor). A file made before site keys has no key file and uses
' the kit's old default key. The folder's permissions are the real control:
' the key also travels in each front end's links.

Public Const PERSONNEL_FILE As String = "Personnel_be.accdb"
Public Const PERSONNEL_FOLDER As String = "Personnel"
Public Const PERSONNEL_KEY_FILE As String = "Personnel.key"
Public Const PERSONNEL_KEY_MIN As Long = 12
Public Const PERSONNEL_KEY_MAX As Long = 20       ' an Access database password's limit
Public Const PERSON_TABLE As String = "kit_person"
Public Const PERSON_APP_TABLE As String = "kit_person_app"
Public Const PERSON_UPN_TABLE As String = "kit_person_upn"

' The demo Personnel file (CreateDemoPersonnelFile): made-up people for
' demos and testing, every one with this app password.
Public Const DEMO_PASSWORD As String = "pass1234"
Public Const DEMO_PEOPLE As Long = 20
Public Const DEMO_ADMIN_USER As String = "demo01"
Public Const DEMO_COMMAND As String = "DEMO DATA (made-up people)"

' PersonnelStatus results.
Public Const PERSONNEL_READY As Long = 0          ' linked and reachable
Public Const PERSONNEL_NOT_SET_UP As Long = 1     ' no file linked or in the default place: offer find or create
Public Const PERSONNEL_LOST As Long = 2           ' linked, but the file can't be reached: offer find, never create
Public Const PERSONNEL_NO_ADMIN As Long = 3       ' linked and reachable, but nobody in it is an active admin (its
                                                  ' people were deleted, modKitPurge): offer creating the first admin

Private Const DEFAULT_KEY As String = "kit-personnel-1"
Private Const MIN_PASSWORD_LEN As Long = 8
' The first character of password_hash for a temporary password an admin set
' (PersonnelResetPassword): the person must change it at their next password
' sign-in. A mark in the hash, not a column, so existing Personnel files
' need no change while other people have them open.
Private Const TEMP_PWD_MARK As String = "!"
Private Const TEMP_PWD_LEN As Long = 12
' No 0/O, 1/l/I: a temporary password is read out or copied by hand.
Private Const TEMP_PWD_CHARS As String = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789"

' The linked Personnel file's password (see the top of the module).
Public Function PersonnelKey() As String
    PersonnelKey = PersonnelKeyFor(PersonnelPath())
End Function

' The password of the Personnel file at path: the first line of the
' Personnel.key beside it, or the old default key when there is none.
Public Function PersonnelKeyFor(ByVal path As String) As String
    Dim kf As String, k As String
    If InStrRev(path, "\") > 0 Then kf = PersonnelKeyPath(path)
    If Len(kf) > 0 Then
        If FileExists(kf) Then
            k = Replace(ReadTextFile(kf), vbCr, vbLf)
            If InStr(k, vbLf) > 0 Then k = Left$(k, InStr(k, vbLf) - 1)
            k = Trim$(k)
        End If
    End If
    If Len(k) = 0 Then k = DEFAULT_KEY
    PersonnelKeyFor = k
End Function

' The key file for the Personnel file at path: Personnel.key in its folder.
Public Function PersonnelKeyPath(ByVal path As String) As String
    PersonnelKeyPath = Left$(path, InStrRev(path, "\")) & PERSONNEL_KEY_FILE
End Function

' A new random site key: PERSONNEL_KEY_MAX hex characters.
Public Function NewPersonnelKey() As String
    NewPersonnelKey = UCase$(RandomHex(PERSONNEL_KEY_MAX \ 2))
End Function

' Why key can't be a site key ("" when it can): its length, or a character
' that would break the file's connection string.
Public Function PersonnelKeyProblem(ByVal key As String) As String
    Dim i As Long, ch As String
    If Len(key) < PERSONNEL_KEY_MIN Or Len(key) > PERSONNEL_KEY_MAX Then
        PersonnelKeyProblem = "Use " & PERSONNEL_KEY_MIN & " to " & PERSONNEL_KEY_MAX & " characters for the site key."
        Exit Function
    End If
    For i = 1 To Len(key)
        ch = Mid$(key, i, 1)
        If AscW(ch) < 33 Or AscW(ch) > 126 Or InStr(";'""[]", ch) > 0 Then
            PersonnelKeyProblem = "Use letters, digits and punctuation for the site key, with no spaces or ; ' "" [ ]."
            Exit Function
        End If
    Next
End Function

' The linked Personnel file, or "" when the tables aren't linked.
Public Function PersonnelPath() As String
    PersonnelPath = KitBackendPath(PERSON_TABLE)
End Function

' Where the Personnel file is looked for when none is linked: a Personnel
' folder beside the folder of appTable's data file, or beside the front end
' when appTable isn't linked.
Public Function PersonnelDefaultPath(ByVal appTable As String) As String
    Dim folder As String, parent As String
    folder = KitBackendPath(appTable)
    If Len(folder) > 0 Then folder = Left$(folder, InStrRev(folder, "\") - 1) Else folder = CurrentProject.Path
    If InStrRev(folder, "\") > 0 Then         ' "S:" (a file at a drive root) has no parent
        parent = Left$(folder, InStrRev(folder, "\") - 1)
        If Not (parent Like "\\*" And InStr(3, parent, "\") = 0) Then folder = parent   ' never up to just "\\server"
    End If
    PersonnelDefaultPath = folder & "\" & PERSONNEL_FOLDER & "\" & PERSONNEL_FILE
End Function

' Whether the Personnel directory can be used (PERSONNEL_READY), and if not
' why. Links the file found in the default place, so a second app finds it
' with no questions.
Public Function PersonnelStatus(ByVal appTable As String) As Long
    Dim def As String
    On Error GoTo Fail
    If Len(PersonnelPath()) > 0 Then
        If KitBackendReachable(PERSON_TABLE) Then PersonnelStatus = ReadyOrNoAdmin(): Exit Function
    End If
    def = PersonnelDefaultPath(appTable)
    If Len(PersonnelFileProblem(def)) = 0 Then
        LinkPersonnelFile def
        PersonnelStatus = ReadyOrNoAdmin()
    ElseIf Len(PersonnelPath()) > 0 Then
        PersonnelStatus = PERSONNEL_LOST
    Else
        PersonnelStatus = PERSONNEL_NOT_SET_UP
    End If
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelStatus", appTable
End Function

' A reachable, linked Personnel file: PERSONNEL_READY, or PERSONNEL_NO_ADMIN
' when nobody in it is an active admin.
Private Function ReadyOrNoAdmin() As Long
    If DCount("*", PERSON_TABLE, "is_admin AND is_active") = 0 Then ReadyOrNoAdmin = PERSONNEL_NO_ADMIN Else ReadyOrNoAdmin = PERSONNEL_READY
End Function

' Why path can't be used as the Personnel file ("" when it can): missing,
' not opening with its key (PersonnelKeyFor), or without the two tables.
Public Function PersonnelFileProblem(ByVal path As String) As String
    Dim db As DAO.Database
    If Not FileExists(path) Then PersonnelFileProblem = "There is no file at " & path & ".": Exit Function
    On Error Resume Next
    Set db = DBEngine.OpenDatabase(path, False, True, ";PWD=" & PersonnelKeyFor(path))
    On Error GoTo 0
    If db Is Nothing Then
        PersonnelFileProblem = "That file can't be opened as the Personnel file. It may be another app's file, or its " & _
            PERSONNEL_KEY_FILE & " may be missing or changed."
        Exit Function
    End If
    If Not (HasTable(db, PERSON_TABLE) And HasTable(db, PERSON_APP_TABLE)) Then
        PersonnelFileProblem = "That file isn't a Personnel file: it has no " & PERSON_TABLE & " table."
    End If
    db.Close
End Function

' Link the Personnel tables to path (replacing any links to another file),
' after checking it with PersonnelFileProblem (error 5 with its words).
Public Sub LinkPersonnelFile(ByVal path As String)
    Dim problem As String, n As Long, d As String, src As String
    KitBusy "Connecting to the Personnel file..."   ' a file on a share can take a while
    On Error GoTo Fail
    problem = PersonnelFileProblem(path)
    If Len(problem) > 0 Then Err.Raise 5, "modPersonnel.LinkPersonnelFile", problem
    KitLinkTables path, Array(PERSON_TABLE, PERSON_APP_TABLE), PersonnelKeyFor(path)
    KitBusyEnd
    Exit Sub
Fail:
    ' KitBusyEnd resets Err, so keep the error and put it back.
    n = Err.Number: d = Err.Description: src = Err.Source
    KitBusyEnd
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modPersonnel.LinkPersonnelFile", path
End Sub

' Create a new, encrypted Personnel file at path with the site key siteKey
' (a new random one when empty), saved in Personnel.key beside it; link it,
' and add the current Windows user as its first person and directory admin,
' with access to appCode. Refuses (error 5) while a Personnel file is
' linked, so nobody makes a second directory by mistake, when path or the
' key file exists, or when the key is unusable (PersonnelKeyProblem).
' Returns the admin's person_id.
Public Function CreatePersonnelFile(ByVal path As String, ByVal appCode As String, ByVal firstName As String, _
        ByVal lastName As String, Optional ByVal rank As String = "", Optional ByVal siteKey As String = "") As Long
    On Error GoTo Fail
    CreatePersonnelFile = MakePersonnelFile(path, appCode, siteKey, False, firstName, lastName, rank)
    Exit Function
Fail:
    RaiseAgain "modPersonnel.CreatePersonnelFile"
End Function

' Create a new Personnel file of DEMO DATA at path, as CreatePersonnelFile
' does (the same refusals and site key), with the DEMO_PEOPLE made-up people
' of DemoPeople in it instead of the current Windows user: each with the
' app password DEMO_PASSWORD (hashed and salted like any other), access to
' appCode, a made-up DoD ID (its UPN) and DEMO_COMMAND as their command, so
' each record says it is demo data (sign-in names demo01 on, too). No Windows account is
' saved for any of them, so nobody is signed in: the sign-in screen comes
' next (PersonnelSignInWhyText "demo"). For demos and testing only.
' Returns the first demo admin's person_id (DEMO_ADMIN_USER).
Public Function CreateDemoPersonnelFile(ByVal path As String, ByVal appCode As String, _
        Optional ByVal siteKey As String = "") As Long
    On Error GoTo Fail
    CreateDemoPersonnelFile = MakePersonnelFile(path, appCode, siteKey, True)
    Exit Function
Fail:
    RaiseAgain "modPersonnel.CreateDemoPersonnelFile"
End Function

' CreatePersonnelFile and CreateDemoPersonnelFile: the file, its key and
' links, then its first admin (the current Windows user) or the demo people.
' Anything that fails takes the half-made file back.
Private Function MakePersonnelFile(ByVal path As String, ByVal appCode As String, ByVal siteKey As String, _
        ByVal demo As Boolean, Optional ByVal firstName As String = "", Optional ByVal lastName As String = "", _
        Optional ByVal rank As String = "") As Long
    Dim db As DAO.Database, n As Long, d As String, src As String, made As Boolean, keyMade As Boolean, v As Variant
    Dim problem As String
    On Error GoTo Fail
    If Len(PersonnelPath()) > 0 Then Err.Raise 5, "modPersonnel", "A Personnel file is already linked: " & PersonnelPath()
    If FileExists(path) Then Err.Raise 5, "modPersonnel", "A file already exists at " & path & "."
    If FileExists(PersonnelKeyPath(path)) Then Err.Raise 5, "modPersonnel", "A key file already exists at " & PersonnelKeyPath(path) & "."
    siteKey = Trim$(siteKey)
    If Len(siteKey) = 0 Then siteKey = NewPersonnelKey()
    problem = PersonnelKeyProblem(siteKey)
    If Len(problem) > 0 Then Err.Raise 5, "modPersonnel", problem
    EnsureFolder Left$(path, InStrRev(path, "\") - 1)
    WriteTextFile PersonnelKeyPath(path), siteKey & vbCrLf
    keyMade = True
    Set db = DBEngine.CreateDatabase(path, dbLangGeneral & ";pwd=" & siteKey)
    made = True
    CreatePersonnelTables db
    db.Close
    Set db = Nothing
    LinkPersonnelFile path
    If demo Then
        MakePersonnelFile = AddDemoPeople(path, appCode)
    Else
        MakePersonnelFile = AddPerson(firstName, lastName, rank, WinAccount(), "", True)
        KitExecute "UPDATE " & PERSON_TABLE & " SET windows_sid=" & SqlOrNull(WinSid()) & " WHERE person_id=" & MakePersonnelFile
        SetPersonApp MakePersonnelFile, appCode, True
    End If
    Exit Function
Fail:
    ' Take back a half-made file and its links, so a second try isn't
    ' refused as "already linked" with no admin in the file.
    n = Err.Number: d = Err.Description: src = Err.Source        ' Close resets Err
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    If made Then
        For Each v In Array(PERSON_APP_TABLE, PERSON_TABLE)
            If StrComp(KitBackendPath(CStr(v)), path, vbTextCompare) = 0 Then CurrentDb.TableDefs.Delete v
        Next
        Kill path
    End If
    If keyMade Then Kill PersonnelKeyPath(path)
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modPersonnel.MakePersonnelFile", path
End Function

' The demo people (CreateDemoPersonnelFile), one "rank|first|last|admin"
' each. Made-up names, chosen to read as samples; a mix of officers, chiefs
' and enlisted, two of them Personnel admins.
Private Function DemoPeople() As Variant
    DemoPeople = Array("CDR|Avery|Sample|1", "YNC|Blake|Example|1", "LCDR|Casey|Testwell|0", "LT|Dana|Mockford|0", _
        "LTJG|Emery|Placeholder|0", "ENS|Finley|Demoson|0", "CWO3|Gray|Fictive|0", "MMCM|Harper|Specimen|0", _
        "ETCS|Indigo|Prototype|0", "STSC|Jordan|Notreal|0", "MM1|Kendall|Madeup|0", "ET1|Logan|Pretendo|0", _
        "YN1|Morgan|Trialman|0", "IT2|Noel|Dummyfield|0", "LS2|Oakley|Mockridge|0", "MM2|Parker|Samplewood|0", _
        "ET3|Quinn|Testa|0", "STS3|Reese|Examplar|0", "SN|Sawyer|Fakeley|0", "FN|Taylor|Pseudonym|0")
End Function

' Add the demo people to the new file at path (just linked) with access to
' appCode; returns the first one's person_id. Sign-in names demo01 to
' demo20; email first.last@demo.invalid (a domain that can never be real);
' DoD IDs 0000000001 to 0000000020 as their UPNs.
Private Function AddDemoPeople(ByVal path As String, ByVal appCode As String) As Long
    Dim v As Variant, f As Variant, i As Long, id As Long, udb As DAO.Database, nm As String
    Dim n As Long, d As String, src As String
    On Error GoTo Fail
    For Each v In DemoPeople()
        i = i + 1
        f = Split(v, "|")
        nm = "demo" & Format$(i, "00")
        id = AddPerson(f(1), f(2), f(0), "", nm, f(3) = "1", "", DEMO_COMMAND, _
            LCase$(f(1) & "." & f(2)) & "@demo.invalid")
        PersonnelSetPassword id, DEMO_PASSWORD
        SetPersonApp id, appCode, True
        If i = 1 Then AddDemoPeople = id
    Next
    Set udb = DBEngine.OpenDatabase(path, False, False, ";PWD=" & PersonnelKeyFor(path))
    i = 0
    For Each v In DemoPeople()
        i = i + 1
        SaveUpn udb, PersonIdFor("demo" & Format$(i, "00")), PersonnelUpnFor(Format$(i, "0000000000"))
    Next
    udb.Close
    Exit Function
Fail:
    ' Close the file first, so MakePersonnelFile can take it back.
    n = Err.Number: d = Err.Description: src = Err.Source
    On Error Resume Next
    If Not udb Is Nothing Then udb.Close
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modPersonnel.AddDemoPeople", path
End Function

' Make the current Windows user the first admin of the linked Personnel file
' after its people were deleted (PERSONNEL_NO_ADMIN), with access to appCode,
' as CreatePersonnelFile does for a new file. Refuses (error 5) while the
' file has an active admin, so nobody can make themselves admin of a
' directory in use. Returns the admin's person_id.
Public Function AddFirstAdmin(ByVal appCode As String, ByVal firstName As String, ByVal lastName As String, _
        Optional ByVal rank As String = "") As Long
    On Error GoTo Fail
    If Len(PersonnelPath()) = 0 Then Err.Raise 5, "modPersonnel", "No Personnel file is linked."
    If ReadyOrNoAdmin() <> PERSONNEL_NO_ADMIN Then Err.Raise 5, "modPersonnel", "The Personnel file already has an admin."
    AddFirstAdmin = AddPerson(firstName, lastName, rank, WinAccount(), "", True)
    KitExecute "UPDATE " & PERSON_TABLE & " SET windows_sid=" & SqlOrNull(WinSid()) & " WHERE person_id=" & AddFirstAdmin
    SetPersonApp AddFirstAdmin, appCode, True
    Exit Function
Fail:
    RaiseAgain "modPersonnel.AddFirstAdmin", firstName & " " & lastName
End Function

' Add a person; returns their person_id. windowsAccount (DOMAIN\user) may be
' ""; username defaults to its user part. Raises error 5 when the sign-in
' name or Windows account is already someone's.
Public Function AddPerson(ByVal firstName As String, ByVal lastName As String, Optional ByVal rank As String = "", _
        Optional ByVal windowsAccount As String = "", Optional ByVal username As String = "", _
        Optional ByVal isAdmin As Boolean = False, Optional ByVal middleInitial As String = "", _
        Optional ByVal commandName As String = "", Optional ByVal email As String = "") As Long
    On Error GoTo Fail
    firstName = Trim$(firstName): lastName = Trim$(lastName)
    If Len(firstName) = 0 Or Len(lastName) = 0 Then Err.Raise 5, "modPersonnel", "A person needs a first and last name."
    If Len(username) = 0 Then username = SignInNameFor(windowsAccount)
    If Len(username) = 0 Then username = LCase$(Left$(firstName, 1) & Replace(lastName, " ", ""))
    If Not IsNull(DLookup("person_id", PERSON_TABLE, "username=" & SqlText(username))) Then
        Err.Raise 5, "modPersonnel", "The sign-in name " & username & " is already used."
    End If
    If Len(windowsAccount) > 0 Then
        If Not IsNull(DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(windowsAccount))) Then
            Err.Raise 5, "modPersonnel", "The Windows account " & windowsAccount & " is already someone's."
        End If
    End If
    KitExecute "INSERT INTO " & PERSON_TABLE & " (windows_account, username, rank, first_name, middle_initial, last_name, " & _
        "display_name, command, email, is_admin, is_active, " & StampColumns() & ") VALUES (" & _
        SqlOrNull(windowsAccount) & ", " & SqlText(username) & ", " & SqlOrNull(rank) & ", " & SqlText(firstName) & ", " & _
        SqlOrNull(Left$(middleInitial, 1)) & ", " & SqlText(lastName) & ", " & _
        SqlText(PersonDisplayName(rank, firstName, lastName)) & ", " & SqlOrNull(commandName) & ", " & SqlOrNull(email) & ", " & _
        IIf(isAdmin, "True", "False") & ", True, " & StampValues() & ")"
    AddPerson = DLookup("person_id", PERSON_TABLE, "username=" & SqlText(username))
    Exit Function
Fail:
    RaiseAgain "modPersonnel.AddPerson", firstName & " " & lastName
End Function

' Give (or take away) a person's access to appCode. Refused (error 5) when
' the signed-in person would take away their own access, which would lock
' the last admin out of the app with no one able to put it back.
Public Sub SetPersonApp(ByVal personId As Long, ByVal appCode As String, ByVal allowed As Boolean)
    Dim crit As String
    On Error GoTo Fail
    If Not allowed And personId = SignedInPersonID() Then
        Err.Raise 5, "modPersonnel", "You can't take away your own access to " & APP_NAME & "."
    End If
    crit = "person_id=" & personId & " AND app_code=" & SqlText(appCode)
    If IsNull(DLookup("person_id", PERSON_APP_TABLE, crit)) Then
        KitExecute "INSERT INTO " & PERSON_APP_TABLE & " (person_id, app_code, is_active, granted_at, granted_by) VALUES (" & _
            personId & ", " & SqlText(appCode) & ", " & IIf(allowed, "True", "False") & ", Now(), " & SqlText(KitUser()) & ")"
    Else
        KitExecute "UPDATE " & PERSON_APP_TABLE & " SET is_active=" & IIf(allowed, "True", "False") & _
            ", granted_at=Now(), granted_by=" & SqlText(KitUser()) & " WHERE " & crit
    End If
    Exit Sub
Fail:
    RaiseAgain "modPersonnel.SetPersonApp", personId & " " & appCode
End Sub

' True when the person is active and may open appCode.
Public Function PersonCanUse(ByVal personId As Long, ByVal appCode As String) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT p.person_id FROM " & PERSON_TABLE & " AS p INNER JOIN " & PERSON_APP_TABLE & _
        " AS a ON p.person_id = a.person_id WHERE p.person_id=" & personId & " AND p.is_active AND a.is_active AND a.app_code=" & _
        SqlText(appCode), dbOpenSnapshot)
    PersonCanUse = Not rs.EOF
    rs.Close
End Function

' "LT J. Smith": rank, first initial, last name.
Public Function PersonDisplayName(ByVal rank As String, ByVal firstName As String, ByVal lastName As String) As String
    PersonDisplayName = Trim$(Trim$(rank) & " " & IIf(Len(Trim$(firstName)) > 0, UCase$(Left$(Trim$(firstName), 1)) & ". ", "") & Trim$(lastName))
End Function

' Sign in the person whose Windows account is running Access
' (PersonnelWindowsPersonId). True when they are active and may open
' appCode; then the session is set and a toast says who. False otherwise.
Public Function PersonnelSignInWindows(ByVal appCode As String, Optional ByVal quiet As Boolean = False) As Boolean
    Dim id As Long
    On Error GoTo Fail
    id = PersonnelWindowsPersonId()
    If id = 0 Then Exit Function
    If Not PersonCanUse(id, appCode) Then Exit Function
    StartSession id, quiet
    PersonnelSignInWindows = True
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelSignInWindows", appCode
End Function

' The person_id saved for the Windows account running Access, active or
' not, or 0: by SID, else by DOMAIN\user (saving the SID, so a renamed
' account still matches).
Public Function PersonnelWindowsPersonId() As Long
    Dim id As Variant, sid As String
    id = Null                                   ' Empty would skip the DOMAIN\user lookup
    sid = WinSid()
    If Len(sid) > 0 Then id = DLookup("person_id", PERSON_TABLE, "windows_sid=" & SqlText(sid))
    If IsNull(id) And Len(WinAccount()) > 0 Then
        id = DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(WinAccount()) & " AND windows_sid Is Null")
        If Not IsNull(id) And Len(sid) > 0 Then
            KitExecute "UPDATE " & PERSON_TABLE & " SET windows_sid=" & SqlText(sid) & " WHERE person_id=" & id
        End If
    End If
    PersonnelWindowsPersonId = Nz(id, 0)
End Function

' Who to ask for a Personnel change: " (LT J. Smith or CDR K. Lee)", the
' names of up to three active directory admins, or "" when there are none.
' For the end of "Ask a Personnel admin".
Public Function PersonnelAdminNames() As String
    Dim rs As DAO.Recordset, names(1 To 3) As String, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT display_name, username FROM " & PERSON_TABLE & _
        " WHERE is_admin AND is_active ORDER BY last_name, first_name", dbOpenSnapshot)
    Do While Not rs.EOF And n < 3
        n = n + 1
        names(n) = Nz(rs!display_name, rs!username)
        rs.MoveNext
    Loop
    rs.Close
    Select Case n
        Case 1: PersonnelAdminNames = " (" & names(1) & ")"
        Case 2: PersonnelAdminNames = " (" & names(1) & " or " & names(2) & ")"
        Case 3: PersonnelAdminNames = " (" & names(1) & ", " & names(2) & " or " & names(3) & ")"
    End Select
End Function

' The login screen's check: "" and the session set when username and
' password match an active person who may open appCode, else what to tell
' the user (never which of the two was wrong). A person with no password
' yet can't sign in here: PersonnelNeedsPassword says so, and
' PersonnelSetPassword sets it.
Public Function PersonnelSignIn(ByVal username As String, ByVal password As String, ByVal appCode As String) As String
    Dim rs As DAO.Recordset, ok As Boolean
    On Error GoTo Fail
    Set rs = CurrentDb.OpenRecordset("SELECT person_id, password_hash FROM " & PERSON_TABLE & _
        " WHERE is_active AND username=" & SqlText(Trim$(username)), dbOpenSnapshot)
    If Not rs.EOF Then ok = VerifyPassword(password, HashOnly(Nz(rs!password_hash, "")))
    If ok Then
        If Not PersonCanUse(rs!person_id, appCode) Then
            PersonnelSignIn = "You don't have access to " & APP_NAME & " yet. Ask a Personnel admin" & PersonnelAdminNames() & " to give you access."
        Else
            StartSession rs!person_id, False
        End If
    Else
        PersonnelSignIn = "That sign-in name and password don't match."
    End If
    rs.Close
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelSignIn", username
End Function

' True when username belongs to an active person who has no app password
' yet (the login screen then asks them to set one).
Public Function PersonnelNeedsPassword(ByVal username As String) As Boolean
    Dim crit As String
    ' Not Nz in the DLookup expression: over a linked table it returns Null
    ' for a missing hash, the same as for no such person.
    crit = "is_active AND username=" & SqlText(Trim$(username))
    If DCount("*", PERSON_TABLE, crit) = 0 Then Exit Function
    PersonnelNeedsPassword = (Len(Nz(DLookup("password_hash", PERSON_TABLE, crit), "")) = 0)
End Function

' The person_id of an active person with sign-in name username, or 0.
Public Function PersonIdFor(ByVal username As String) As Long
    PersonIdFor = Nz(DLookup("person_id", PERSON_TABLE, "is_active AND username=" & SqlText(Trim$(username))), 0)
End Function

' True when the Windows account running Access is username's saved account,
' which proves who they are, so they may choose their first app password
' themselves. Anyone else needs a Personnel admin to set it.
Public Function PersonnelCanSetFirstPassword(ByVal username As String) As Boolean
    Dim rs As DAO.Recordset
    If Len(WinAccount()) = 0 Or Not PersonnelNeedsPassword(username) Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT windows_account, windows_sid FROM " & PERSON_TABLE & _
        " WHERE username=" & SqlText(Trim$(username)), dbOpenSnapshot)
    ' The same proof as PersonnelSignInWindows: a saved SID must match (the
    ' DOMAIN\user name may since belong to someone else); the account name
    ' counts only while no SID is saved.
    If Not IsNull(rs!windows_sid) Then
        PersonnelCanSetFirstPassword = (Len(WinSid()) > 0 And StrComp(rs!windows_sid, WinSid(), vbTextCompare) = 0)
    Else
        PersonnelCanSetFirstPassword = (StrComp(Nz(rs!windows_account, ""), WinAccount(), vbTextCompare) = 0)
    End If
    rs.Close
End Function

' True when the Windows account running Access is saved for any person.
Public Function WindowsAccountKnown() As Boolean
    If Len(WinAccount()) > 0 Then WindowsAccountKnown = Not IsNull(DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(WinAccount())))
End Function

' Set a person's app password (salted hash). Returns "" or why the password
' is refused. temporary marks it as one the person must change when they
' next sign in with it (an admin set it, so the admin knows it).
Public Function PersonnelSetPassword(ByVal personId As Long, ByVal password As String, _
        Optional ByVal temporary As Boolean = False) As String
    If Len(password) < MIN_PASSWORD_LEN Then
        PersonnelSetPassword = "Use at least " & MIN_PASSWORD_LEN & " characters."
        Exit Function
    End If
    KitExecute "UPDATE " & PERSON_TABLE & " SET password_hash=" & SqlText(IIf(temporary, TEMP_PWD_MARK, "") & HashPassword(password)) & _
        ", " & StampSql() & " WHERE person_id=" & personId
End Function

' An admin's password reset: a new random temporary password for personId,
' which they must change when they sign in with it. Returns it, to show the
' admin once; it is never saved, only its hash. Refused (error 5) unless the
' signed-in person is an active admin, and for the admin's own password
' (KitPersonnelChangePassword is for that). Written to the Windows event log.
Public Function PersonnelResetPassword(ByVal personId As Long) As String
    Dim pwd As String, who As String
    On Error GoTo Fail
    PersonnelRequireAdmin
    If personId = SignedInPersonID() Then Err.Raise 5, "modPersonnel", "Use Change my password for your own password."
    who = Nz(DLookup("username", PERSON_TABLE, "person_id=" & personId), "")
    If Len(who) = 0 Then Err.Raise 5, "modPersonnel", "That person isn't in the Personnel file."
    pwd = TemporaryPassword()
    PersonnelSetPassword personId, pwd, True
    KitEventLog "Personnel: " & KitUser() & " reset the app password of " & who & "."
    PersonnelResetPassword = pwd
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelResetPassword", CStr(personId)
End Function

' True when personId's app password is a temporary one an admin set, which
' they must change at their next password sign-in.
Public Function PersonnelPasswordIsTemporary(ByVal personId As Long) As Boolean
    PersonnelPasswordIsTemporary = (Left$(Nz(DLookup("password_hash", PERSON_TABLE, "person_id=" & personId), ""), 1) = TEMP_PWD_MARK)
End Function

' True when personId has an app password (temporary or not).
Public Function PersonnelHasPassword(ByVal personId As Long) As Boolean
    PersonnelHasPassword = (Len(Nz(DLookup("password_hash", PERSON_TABLE, "person_id=" & personId), "")) > 0)
End Function

' A person changing their own app password: "" when done, else why not
' (the current password is wrong, the new one too short or the same). With
' no password yet, currentPassword is ignored: being signed in proves who
' they are. The new password is never temporary.
Public Function PersonnelChangePassword(ByVal personId As Long, ByVal currentPassword As String, ByVal newPassword As String) As String
    Dim stored As String
    stored = Nz(DLookup("password_hash", PERSON_TABLE, "person_id=" & personId), "")
    If Len(stored) > 0 Then
        If Not VerifyPassword(currentPassword, HashOnly(stored)) Then
            PersonnelChangePassword = "Your current password isn't right.": Exit Function
        End If
        If StrComp(currentPassword, newPassword, vbBinaryCompare) = 0 Then
            PersonnelChangePassword = "Choose a password different from the current one.": Exit Function
        End If
    End If
    PersonnelChangePassword = PersonnelSetPassword(personId, newPassword, False)
End Function

' Save the current Windows account to the signed-in person's record ("Use
' my Windows sign-in next time"). Refused (error 5) when the account is
' already someone else's.
Public Sub PersonnelSaveWindowsAccount()
    Dim other As Variant
    If SignedInPersonID() = 0 Then Err.Raise 5, "modPersonnel", "Nobody is signed in."
    other = DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(WinAccount()) & " AND person_id<>" & SignedInPersonID())
    If Not IsNull(other) Then Err.Raise 5, "modPersonnel", "This Windows account is already saved for someone else."
    KitExecute "UPDATE " & PERSON_TABLE & " SET windows_account=" & SqlText(WinAccount()) & ", windows_sid=" & _
        SqlOrNull(WinSid()) & ", " & StampSql() & " WHERE person_id=" & SignedInPersonID()
End Sub

' Sign out: clear the session. The app then shows its login screen (not a
' Windows sign-in), so another person can sign in.
Public Sub PersonnelSignOut()
    Dim v As Variant
    On Error Resume Next                        ' one that was never set
    For Each v In Array("PersonID", "Username", "PersonName", "PersonIsAdmin")
        TempVars.Remove v
    Next
End Sub

' The signed-in person's person_id, or 0.
Public Function SignedInPersonID() As Long
    On Error Resume Next                        ' TempVars may be unset
    SignedInPersonID = Nz(TempVars("PersonID"), 0)
End Function

' True when the signed-in person is a directory admin.
Public Function SignedInIsAdmin() As Boolean
    On Error Resume Next
    SignedInIsAdmin = Nz(TempVars("PersonIsAdmin"), False)
End Function

' True when the signed-in person is an active directory admin in the
' Personnel file now: another admin may have changed them since they
' signed in. Updates PersonIsAdmin to match.
Public Function PersonnelIsAdminNow() As Boolean
    If SignedInPersonID() = 0 Then Exit Function
    PersonnelIsAdminNow = (DCount("*", PERSON_TABLE, "is_admin AND is_active AND person_id=" & SignedInPersonID()) > 0)
    TempVars("PersonIsAdmin") = PersonnelIsAdminNow
End Function

' Raise error 5 unless the signed-in person is an active admin now
' (PersonnelIsAdminNow). Every change to other people goes through this.
Public Sub PersonnelRequireAdmin()
    If Not PersonnelIsAdminNow() Then Err.Raise 5, "modPersonnel", "Only an active Personnel admin can add or change people."
End Sub

' ---------------------------------------------------------------------------
' Directory admin (the Personnel page, dlgPersonnel and dlgPerson)
' ---------------------------------------------------------------------------

' Change a person: AddPerson's fields, plus active. A changed Windows account
' clears the saved SID, which belonged to the old account. Refused (error 5)
' when the sign-in name or Windows account is someone else's, when the
' signed-in admin would take away their own admin or make themselves
' inactive, or when no active admin would be left.
Public Sub UpdatePerson(ByVal personId As Long, ByVal firstName As String, ByVal lastName As String, _
        ByVal rank As String, ByVal windowsAccount As String, ByVal username As String, _
        ByVal isAdmin As Boolean, ByVal isActive As Boolean, Optional ByVal middleInitial As String = "", _
        Optional ByVal commandName As String = "", Optional ByVal email As String = "")
    Dim oldAccount As String, wasAdmin As Boolean, wasActive As Boolean
    On Error GoTo Fail
    PersonnelRequireAdmin
    firstName = Trim$(firstName): lastName = Trim$(lastName)
    username = Trim$(username): windowsAccount = Trim$(windowsAccount)
    If Len(firstName) = 0 Or Len(lastName) = 0 Then Err.Raise 5, "modPersonnel", "A person needs a first and last name."
    If Len(username) = 0 Then Err.Raise 5, "modPersonnel", "A person needs a sign-in name."
    If Not IsNull(DLookup("person_id", PERSON_TABLE, "username=" & SqlText(username) & " AND person_id<>" & personId)) Then
        Err.Raise 5, "modPersonnel", "The sign-in name " & username & " is already used."
    End If
    If Len(windowsAccount) > 0 Then
        If Not IsNull(DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(windowsAccount) & " AND person_id<>" & personId)) Then
            Err.Raise 5, "modPersonnel", "The Windows account " & windowsAccount & " is already someone's."
        End If
    End If
    If personId = SignedInPersonID() Then
        If Not isActive Then Err.Raise 5, "modPersonnel", "You can't make yourself inactive."
        If SignedInIsAdmin() And Not isAdmin Then Err.Raise 5, "modPersonnel", "You can't take away your own admin access."
    End If
    If Not (isAdmin And isActive) Then
        If DCount("*", PERSON_TABLE, "is_admin AND is_active AND person_id<>" & personId) = 0 Then
            Err.Raise 5, "modPersonnel", "Personnel needs at least one active admin. Make someone else an admin first."
        End If
    End If
    oldAccount = Nz(DLookup("windows_account", PERSON_TABLE, "person_id=" & personId), "")
    wasAdmin = Nz(DLookup("is_admin", PERSON_TABLE, "person_id=" & personId), False)
    wasActive = Nz(DLookup("is_active", PERSON_TABLE, "person_id=" & personId), False)
    KitExecute "UPDATE " & PERSON_TABLE & " SET username=" & SqlText(username) & ", rank=" & SqlOrNull(rank) & _
        ", first_name=" & SqlText(firstName) & ", middle_initial=" & SqlOrNull(Left$(Trim$(middleInitial), 1)) & _
        ", last_name=" & SqlText(lastName) & ", display_name=" & SqlText(PersonDisplayName(rank, firstName, lastName)) & _
        ", command=" & SqlOrNull(commandName) & ", email=" & SqlOrNull(email) & _
        ", windows_account=" & SqlOrNull(windowsAccount) & _
        IIf(StrComp(oldAccount, windowsAccount, vbTextCompare) <> 0, ", windows_sid=Null", "") & _
        ", is_admin=" & IIf(isAdmin, "True", "False") & ", is_active=" & IIf(isActive, "True", "False") & _
        ", " & StampSql() & " WHERE person_id=" & personId
    ' Two admins demoting each other from two front ends can both pass the
    ' check above. Look again now the change is saved, and put this row's
    ' admin and active back if nobody is left: at worst both keep admin.
    If Not (isAdmin And isActive) Then
        If DCount("*", PERSON_TABLE, "is_admin AND is_active") = 0 Then
            KitExecute "UPDATE " & PERSON_TABLE & " SET is_admin=" & IIf(wasAdmin, "True", "False") & _
                ", is_active=" & IIf(wasActive, "True", "False") & " WHERE person_id=" & personId
            Err.Raise 5, "modPersonnel", "Personnel needs at least one active admin. Make someone else an admin first."
        End If
    End If
    Exit Sub
Fail:
    RaiseAgain "modPersonnel.UpdatePerson", CStr(personId)
End Sub

' Row source for the Personnel page's list: everyone whose name, rank,
' sign-in name or Windows account contains term, by last name. Columns:
' person_id (hidden), Name, Sign-in name, Windows account, whether they may
' open appCode, and Admin / Active / Inactive.
Public Function PersonnelPeopleSql(ByVal term As String, ByVal appCode As String) As String
    Dim where As String
    term = Trim$(term)
    If Len(term) > 0 Then
        where = " WHERE (p.last_name & ', ' & p.first_name & ' ' & Nz(p.rank,'') & ' ' & p.username & ' ' & " & _
            "Nz(p.windows_account,'')) Like " & SqlText("*" & LikeText(term) & "*")
    End If
    PersonnelPeopleSql = "SELECT p.person_id, p.last_name & ', ' & p.first_name & IIf(p.rank Is Null, '', ' (' & p.rank & ')') AS [Name], " & _
        "p.username AS [Sign-in name], p.windows_account AS [Windows account], " & _
        "IIf(a.is_active, 'Yes', 'No') AS [Can open], " & _
        "IIf(Not p.is_active, 'Inactive', IIf(p.is_admin, 'Admin', 'Active')) AS [Status] " & _
        "FROM " & PERSON_TABLE & " AS p LEFT JOIN (SELECT person_id, is_active FROM " & PERSON_APP_TABLE & _
        " WHERE app_code=" & SqlText(appCode) & ") AS a ON p.person_id = a.person_id" & where & _
        " ORDER BY p.last_name, p.first_name"
End Function

'Import and export (the Personnel page's Import... and Export...)
'
' Import reads a CSV or Excel file (.csv, .xlsx, .xls; first row is the
' headers), including a contacts list exported from Outlook. Headers are
' matched in any order and case, with spaces, underscores and hyphens
' ignored; each field takes the first header it finds:
'   First name     First name, Given name
'   Last name      Last name, Surname, Family name
'                  (or both from Display name, Name, Full name or E-mail
'                  display name: "SMITH, JOHN A" or "John A Smith")
'   Middle initial Middle initial, MI, or Middle name's first letter
'   Rank           Rank, Pay grade, Grade, or Title (not Mr., Mrs., Ms., Dr.)
'   Command        Command, Unit, Company, Organization
'   Email          Email, E-mail address, Mail, SMTP address
'   Sign-in name   Sign-in name, Username
'   Windows        Windows account (DOMAIN\user), Windows SID or SID
'   UPN            User principal name or UPN, or a DoD ID or EDIPI
'                  (which gives <DoD ID>.A@smil.mil)
' A row with no Windows SID is looked up in Active Directory
' (modPersonnelAD.PersonnelFindAccount) by its UPN, else its Windows
' account, else its email; what is found fills its Windows account and
' SID, so the person is signed in by Windows the first time they open an
' app. The UPN is kept in kit_person_upn (admins only, never exported or
' logged) so the account can be found again later.
'
' A row that matches someone already in the file (same SID, Windows account,
' saved UPN, only person with that email, or sign-in name) fills in what
' that person is missing (never changes what is there, their name, admin
' or active),
' and they are given access to appCode unless an admin already decided it.
' Anyone else is added with access to appCode. Returns what happened, for a
' message. Admins only.
Public Function PersonnelImport(ByVal path As String, ByVal appCode As String) As String
    Dim rows As Collection, row As Variant, n As Long, problem As String, outcome As String, lookedUp As String
    Dim added As Long, updated As Long, same As Long, linked As Long, skipped As Long, notFound As Long
    Dim why As String, missing As String, msg As String
    Dim num As Long, d As String, src As String, busy As Boolean, udb As DAO.Database
    On Error GoTo Fail
    PersonnelRequireAdmin
    KitBusy "Importing people..."
    busy = True
    Set rows = ImportRows(path)
    Set udb = UpnDb()
    For Each row In rows
        n = n + 1
        problem = ImportPerson(row, appCode, udb, outcome, lookedUp)
        If Len(problem) = 0 Then
            Select Case outcome
                Case "added": added = added + 1
                Case "updated": updated = updated + 1
                Case Else: same = same + 1
            End Select
            If lookedUp = "found" Then linked = linked + 1
            If lookedUp = "missing" Then
                notFound = notFound + 1
                If notFound <= 5 Then missing = missing & vbCrLf & "- Row " & (n + 1)
            End If
        Else
            skipped = skipped + 1
            If skipped <= 5 Then why = why & vbCrLf & "- Row " & (n + 1) & ": " & problem   ' row 1 is the headers
        End If
    Next
    udb.Close
    Set udb = Nothing
    KitBusyEnd
    busy = False
    If n = 0 Then PersonnelImport = "That file has no people in it.": Exit Function
    msg = "Added " & added & IIf(added = 1, " person", " people") & "."
    If updated > 0 Then msg = msg & " Filled in details for " & updated & " already in Personnel."
    If same > 0 Then msg = msg & " " & same & " already up to date."
    If linked > 0 Then msg = msg & " Linked " & linked & " to their Windows account."
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Skipped " & skipped & ":" & why
        If skipped > 5 Then msg = msg & vbCrLf & "- and " & (skipped - 5) & " more"
    End If
    If notFound > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Not found in Active Directory, so they sign in with an app password until " & _
            "you add their Windows account: " & notFound & missing
        If notFound > 5 Then msg = msg & vbCrLf & "- and " & (notFound - 5) & " more"
    End If
    KitEventLog "Personnel: " & KitUser() & " imported " & path & " (added " & added & ", updated " & updated & _
        ", skipped " & skipped & ")."
    PersonnelImport = msg
    Exit Function
Fail:
    ' KitBusyEnd resets Err, so keep the error and put it back.
    num = Err.Number: d = Err.Description: src = Err.Source
    On Error Resume Next
    If Not udb Is Nothing Then udb.Close
    On Error GoTo 0
    If busy Then KitBusyEnd
    Err.Number = num: Err.Description = d: Err.Source = src
    RaiseAgain "modPersonnel.PersonnelImport", path
End Function

' Write everyone in the Personnel file to a CSV at path, in the headers
' PersonnelImport reads, so the file can be imported into another Personnel
' file. Adds Personnel admin, Active and Apps (the app codes they may open)
' for the reader; Import doesn't use those. No password data and no UPN
' (kit_person_upn holds DoD IDs and stays in the file). Admins only;
' written to the Windows event log. Returns how many people.
Public Function PersonnelExport(ByVal path As String) As Long
    Dim rs As DAO.Recordset, ra As DAO.Recordset, sb As String, apps As String, n As Long
    On Error GoTo Fail
    PersonnelRequireAdmin
    sb = CsvLine("First name", "Middle initial", "Last name", "Rank", "Command", "Email", "Sign-in name", _
        "Windows account", "Windows SID", "Personnel admin", "Active", "Apps")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & PERSON_TABLE & " ORDER BY last_name, first_name", dbOpenSnapshot)
    Do While Not rs.EOF
        apps = ""
        Set ra = CurrentDb.OpenRecordset("SELECT app_code FROM " & PERSON_APP_TABLE & " WHERE is_active AND person_id=" & _
            rs!person_id & " ORDER BY app_code", dbOpenSnapshot)
        Do While Not ra.EOF
            apps = apps & IIf(Len(apps) > 0, " ", "") & ra!app_code
            ra.MoveNext
        Loop
        ra.Close
        sb = sb & CsvLine(rs!first_name, rs!middle_initial, rs!last_name, rs!rank, rs!command, rs!email, rs!username, _
            rs!windows_account, rs!windows_sid, IIf(rs!is_admin, "Yes", "No"), IIf(rs!is_active, "Yes", "No"), apps)
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    WriteTextFile path, sb
    KitEventLog "Personnel: " & KitUser() & " exported " & n & " people to " & path & "."
    PersonnelExport = n
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelExport", path
End Function

' The file's rows as dictionaries keyed by header with spaces, underscores
' and hyphens taken out ("firstname"). Error 5 for a file type it can't read.
Private Function ImportRows(ByVal path As String) As Collection
    Const TMP As String = "zz_personnel_import"
    Dim out As New Collection, d As Object, src As Variant, k As Variant, ext As String
    Dim rs As DAO.Recordset, fld As DAO.Field
    On Error GoTo Fail
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    Select Case ext
    Case "csv"
        For Each src In ReadCsvDicts(path)
            Set d = CreateObject("Scripting.Dictionary")
            For Each k In src.Keys
                d(HeaderKey(k)) = Trim$(Nz(src(k), ""))
            Next
            out.Add d
        Next
    Case "xlsx", "xls"
        On Error Resume Next                    ' left from an earlier import that failed
        CurrentDb.TableDefs.Delete TMP
        On Error GoTo Fail
        DoCmd.TransferSpreadsheet acImport, IIf(ext = "xls", acSpreadsheetTypeExcel8, acSpreadsheetTypeExcel12Xml), TMP, path, True
        Set rs = CurrentDb.OpenRecordset(TMP, dbOpenSnapshot)
        Do Until rs.EOF
            Set d = CreateObject("Scripting.Dictionary")
            For Each fld In rs.Fields
                d(HeaderKey(fld.Name)) = Trim$(Nz(fld.Value, ""))
            Next
            out.Add d
            rs.MoveNext
        Loop
        rs.Close
        CurrentDb.TableDefs.Delete TMP
    Case Else
        Err.Raise 5, "modPersonnel", "Import reads .csv, .xlsx or .xls files."
    End Select
    Set ImportRows = out
    Exit Function
Fail:
    RaiseAgain "modPersonnel.ImportRows", path
End Function

' Add or fill in one imported row (see PersonnelImport). "" when done, with
' outcome "added", "updated" or "same"; else why it was skipped.
' lookedUp is "found" when Active Directory gave the Windows account,
' "missing" when the row had a UPN, DoD ID or Windows account it didn't
' find, else "". udb is the Personnel file with kit_person_upn (UpnDb).
Private Function ImportPerson(ByVal d As Object, ByVal appCode As String, udb As DAO.Database, _
        outcome As String, lookedUp As String) As String
    Dim first As String, last As String, mi As String, rank As String, cmd As String, email As String
    Dim user As String, acct As String, sid As String, upn As String, t As String, foundAcct As String, foundSid As String
    Dim id As Long, other As Variant, sets As String, upnKeep As String, upnId As Long, upnNew As Boolean
    Dim rs As DAO.Recordset
    On Error GoTo Fail
    outcome = "": lookedUp = ""
    first = ImportValue(d, "firstname", "givenname")
    last = ImportValue(d, "lastname", "surname", "familyname")
    mi = Left$(ImportValue(d, "middleinitial", "mi", "middlename"), 1)
    If Len(first) = 0 Or Len(last) = 0 Then
        SplitDisplayName ImportValue(d, "displayname", "name", "fullname", "emaildisplayname"), first, last, mi
    End If
    rank = ImportValue(d, "rank", "paygrade", "grade")
    If Len(rank) = 0 Then
        t = ImportValue(d, "title")
        Select Case LCase$(Replace(t, ".", ""))
            Case "mr", "mrs", "ms", "miss", "mx", "dr", "prof", "sir", "madam": t = ""
        End Select
        If Len(t) <= 20 Then rank = t
    End If
    cmd = ImportValue(d, "command", "unit", "company", "organization")
    email = ImportValue(d, "email", "emailaddress", "mail", "smtpaddress")
    user = ImportValue(d, "signinname", "username")
    acct = ImportValue(d, "windowsaccount", "account")
    If InStr(acct, "\") = 0 Then acct = ""
    sid = UCase$(ImportValue(d, "windowssid", "sid", "objectsid"))
    If Not sid Like "S-1-*" Then sid = ""
    upn = ImportValue(d, "userprincipalname", "upn")
    If Len(upn) = 0 Then upn = PersonnelUpnFor(ImportValue(d, "dodid", "edipi", "dodidnumber"))
    ' Kept only when it is a UPN; anything else in the column (DOMAIN\user) is just looked up.
    If Len(upn) > 0 And Len(UpnProblem(upn)) = 0 Then upnKeep = upn
    upnId = UpnOwner(udb, upnKeep)

    ' Active Directory: the Windows account and SID, unless the row has the SID.
    If Len(sid) = 0 Then
        t = upn
        If Len(t) = 0 Then t = acct
        If Len(t) = 0 Then t = email
        If Len(t) > 0 Then
            If PersonnelFindAccount(t, foundAcct, foundSid) Then
                acct = foundAcct: sid = foundSid: lookedUp = "found"
            ElseIf Len(upn) > 0 Or Len(acct) > 0 Then
                lookedUp = "missing"        ' an email alone not found is no surprise
            End If
        End If
    End If

    ' Someone already here: by SID, Windows account, saved UPN, the one person with this email, or sign-in name.
    If Len(sid) > 0 Then id = Nz(DLookup("person_id", PERSON_TABLE, "windows_sid=" & SqlText(sid)), 0)
    If id = 0 And Len(acct) > 0 Then id = Nz(DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(acct)), 0)
    If id = 0 Then id = upnId
    If id = 0 And Len(email) > 0 Then
        If DCount("*", PERSON_TABLE, "email=" & SqlText(email)) = 1 Then id = DLookup("person_id", PERSON_TABLE, "email=" & SqlText(email))
    End If
    If id = 0 And Len(user) > 0 Then id = Nz(DLookup("person_id", PERSON_TABLE, "username=" & SqlText(user)), 0)

    If id = 0 Then
        If Len(sid) > 0 And Len(acct) = 0 Then Err.Raise 5, "modPersonnel", "It has a Windows SID but no Windows account."
        id = AddPerson(first, last, rank, acct, user, False, mi, cmd, email)
        If Len(sid) > 0 Then KitExecute "UPDATE " & PERSON_TABLE & " SET windows_sid=" & SqlText(sid) & " WHERE person_id=" & id
        If Len(upnKeep) > 0 Then SaveUpn udb, id, upnKeep
        SetPersonApp id, appCode, True
        outcome = "added"
        Exit Function
    End If

    ' Fill in only what they are missing.
    If Len(acct) > 0 Then
        other = DLookup("person_id", PERSON_TABLE, "windows_account=" & SqlText(acct) & " AND person_id<>" & id)
        If Not IsNull(other) Then Err.Raise 5, "modPersonnel", "The Windows account " & acct & " is already someone else's."
    End If
    If Len(sid) > 0 Then
        other = DLookup("person_id", PERSON_TABLE, "windows_sid=" & SqlText(sid) & " AND person_id<>" & id)
        If Not IsNull(other) Then Err.Raise 5, "modPersonnel", "That Windows account is already someone else's."
    End If
    If upnId <> 0 And upnId <> id Then Err.Raise 5, "modPersonnel", "Its user principal name is already someone else's."
    ' Their UPN, unless they have one saved already.
    If Len(upnKeep) > 0 And upnId = 0 Then upnNew = (Len(UpnOf(udb, id)) = 0)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & PERSON_TABLE & " WHERE person_id=" & id, dbOpenSnapshot)
    If Len(acct) > 0 And IsNull(rs!windows_account) Then
        sets = sets & ", windows_account=" & SqlText(acct)
        If Len(sid) > 0 And IsNull(rs!windows_sid) Then sets = sets & ", windows_sid=" & SqlText(sid)
    ElseIf Len(sid) > 0 And IsNull(rs!windows_sid) Then
        ' Only the account's own SID: the saved DOMAIN\user must be the one AD found.
        If StrComp(Nz(rs!windows_account, ""), acct, vbTextCompare) = 0 Then sets = sets & ", windows_sid=" & SqlText(sid)
    End If
    If Len(email) > 0 And IsNull(rs!email) Then sets = sets & ", email=" & SqlText(email)
    If Len(rank) > 0 And IsNull(rs!rank) Then
        sets = sets & ", rank=" & SqlText(rank) & ", display_name=" & SqlText(PersonDisplayName(rank, rs!first_name, rs!last_name))
    End If
    If Len(cmd) > 0 And IsNull(rs!command) Then sets = sets & ", command=" & SqlText(cmd)
    If Len(mi) > 0 And IsNull(rs!middle_initial) Then sets = sets & ", middle_initial=" & SqlText(mi)
    rs.Close
    If upnNew Then SaveUpn udb, id, upnKeep
    If Len(sets) > 0 Then
        KitExecute "UPDATE " & PERSON_TABLE & " SET " & Mid$(sets, 3) & ", " & StampSql() & " WHERE person_id=" & id
        outcome = "updated"
    ElseIf upnNew Then
        outcome = "updated"
    Else
        outcome = "same"
        If lookedUp = "found" Then lookedUp = ""        ' already linked: nothing new
    End If
    ' Access to the app, unless an admin already gave or took it away.
    If IsNull(DLookup("person_id", PERSON_APP_TABLE, "person_id=" & id & " AND app_code=" & SqlText(appCode))) Then
        SetPersonApp id, appCode, True
        If outcome = "same" Then outcome = "updated"
    End If
    Exit Function
Fail:
    If Err.Number <> 5 Then RaiseAgain "modPersonnel.ImportPerson"
    ImportPerson = Err.Description
End Function

' The first of keys the row has a value for, or "".
Private Function ImportValue(ByVal d As Object, ParamArray keys() As Variant) As String
    Dim k As Variant
    For Each k In keys
        If d.Exists(k) Then
            If Len(d(k)) > 0 Then ImportValue = d(k): Exit Function
        End If
    Next
End Function

' First, last name and middle initial from a display name: "SMITH, JOHN A"
' or "John A Smith", with anything in brackets dropped ("(USN)"). Fills
' only the parts still empty.
Private Sub SplitDisplayName(ByVal s As String, first As String, last As String, mi As String)
    Dim p As Long, q As Long, toks() As String, rest As String, f As String, l As String, m As String
    Do
        p = InStr(s, "("): If p = 0 Then Exit Do
        q = InStr(p, s, ")"): If q = 0 Then q = Len(s)
        s = Left$(s, p - 1) & Mid$(s, q + 1)
    Loop
    s = Trim$(Replace(s, "  ", " "))
    If Len(s) = 0 Then Exit Sub
    If InStr(s, ",") > 0 Then
        l = Trim$(Left$(s, InStr(s, ",") - 1))
        rest = Trim$(Mid$(s, InStr(s, ",") + 1))
        If Len(rest) > 0 Then
            toks = Split(rest, " ")
            f = toks(0)
            If UBound(toks) >= 1 Then m = InitialOf(toks(1))
        End If
    Else
        toks = Split(s, " ")
        If UBound(toks) < 1 Then Exit Sub
        f = toks(0): l = toks(UBound(toks))
        If UBound(toks) >= 2 Then m = InitialOf(toks(1))
    End If
    If Len(first) = 0 Then first = f
    If Len(last) = 0 Then last = l
    If Len(mi) = 0 Then mi = m
End Sub

' tok as a middle initial when it is one letter ("A" or "A."), else "".
Private Function InitialOf(ByVal tok As String) As String
    If Len(Replace(tok, ".", "")) = 1 Then InitialOf = Left$(tok, 1)
End Function

Private Function HeaderKey(ByVal s As String) As String
    HeaderKey = LCase$(Replace(Replace(Replace(Trim$(s), " ", ""), "_", ""), "-", ""))
End Function

' s with Like's wildcards made literal.
Private Function LikeText(ByVal s As String) As String
    LikeText = Replace(Replace(Replace(Replace(s, "[", "[[]"), "*", "[*]"), "?", "[?]"), "#", "[#]")
End Function

' ---------------------------------------------------------------------------
' User principal names (kit_person_upn): admins only, never exported or logged
' ---------------------------------------------------------------------------

' The person's saved user principal name, or "". Admins only.
Public Function PersonnelUpn(ByVal personId As Long) As String
    Dim udb As DAO.Database, n As Long, d As String, src As String
    On Error GoTo Fail
    PersonnelRequireAdmin
    Set udb = UpnDb()
    PersonnelUpn = UpnOf(udb, personId)
    udb.Close
    Exit Function
Fail:
    n = Err.Number: d = Err.Description: src = Err.Source
    On Error Resume Next
    If Not udb Is Nothing Then udb.Close
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modPersonnel.PersonnelUpn", CStr(personId)
End Function

' Why upn can't be saved as personId's user principal name ("" when it
' can): not name@domain, or someone else's. A DoD ID (10 digits) counts as
' its UPN (PersonnelUpnFor); "" is fine (no UPN). Admins only.
Public Function PersonnelUpnProblem(ByVal personId As Long, ByVal upn As String) As String
    Dim udb As DAO.Database, owner As Long, n As Long, d As String, src As String
    On Error GoTo Fail
    PersonnelRequireAdmin
    upn = UpnFromInput(upn)
    If Len(upn) = 0 Then Exit Function
    PersonnelUpnProblem = UpnProblem(upn)
    If Len(PersonnelUpnProblem) > 0 Then Exit Function
    Set udb = UpnDb()
    owner = UpnOwner(udb, upn)
    udb.Close
    If owner <> 0 And owner <> personId Then PersonnelUpnProblem = "That user principal name is already someone else's."
    Exit Function
Fail:
    n = Err.Number: d = Err.Description: src = Err.Source
    On Error Resume Next
    If Not udb Is Nothing Then udb.Close
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modPersonnel.PersonnelUpnProblem", CStr(personId)
End Function

' Save upn (a UPN, or a DoD ID for its UPN) as the person's user principal
' name, replacing theirs; "" removes it. "" when saved, else why not
' (PersonnelUpnProblem), with nothing changed. Error 5 when personId isn't
' in Personnel, so no UPN is held for nobody. Admins only; not logged.
Public Function PersonnelSetUpn(ByVal personId As Long, ByVal upn As String) As String
    Dim udb As DAO.Database, n As Long, d As String, src As String
    On Error GoTo Fail
    PersonnelSetUpn = PersonnelUpnProblem(personId, upn)       ' checks admin
    If Len(PersonnelSetUpn) > 0 Then Exit Function
    If DCount("*", PERSON_TABLE, "person_id=" & personId) = 0 Then Err.Raise 5, "modPersonnel", "There is no such person."
    upn = UpnFromInput(upn)
    Set udb = UpnDb()
    If StrComp(UpnOf(udb, personId), upn, vbBinaryCompare) <> 0 Then SaveUpn udb, personId, upn
    udb.Close
    Exit Function
Fail:
    n = Err.Number: d = Err.Description: src = Err.Source
    On Error Resume Next
    If Not udb Is Nothing Then udb.Close
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modPersonnel.PersonnelSetUpn", CStr(personId)
End Function

' Link a person who has no Windows account yet to the one Active Directory
' finds for their saved UPN (modPersonnelAD.PersonnelFindAccount). True when
' it was found and saved; False when they have an account already, no UPN,
' it isn't found, or it is someone else's. Admins only.
Public Function PersonnelLinkFromUpn(ByVal personId As Long) As Boolean
    Dim upn As String, acct As String, sid As String
    On Error GoTo Fail
    upn = PersonnelUpn(personId)                                ' checks admin
    If Len(upn) = 0 Then Exit Function
    If Len(Nz(DLookup("windows_account", PERSON_TABLE, "person_id=" & personId), "")) > 0 Then Exit Function
    If Not PersonnelFindAccount(upn, acct, sid) Then Exit Function
    If Not IsNull(DLookup("person_id", PERSON_TABLE, "(windows_account=" & SqlText(acct) & " OR windows_sid=" & _
        SqlText(sid) & ") AND person_id<>" & personId)) Then Exit Function
    KitExecute "UPDATE " & PERSON_TABLE & " SET windows_account=" & SqlText(acct) & ", windows_sid=" & SqlText(sid) & _
        ", " & StampSql() & " WHERE person_id=" & personId
    PersonnelLinkFromUpn = True
    Exit Function
Fail:
    RaiseAgain "modPersonnel.PersonnelLinkFromUpn", CStr(personId)
End Function

' The Personnel file's tables, children first, for emptying it
' (modKitPurge): kit_person_upn too when the file has it.
Public Function PersonnelDataTables() As Variant
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(PersonnelPath(), False, True, ";PWD=" & PersonnelKey())
    If HasTable(db, PERSON_UPN_TABLE) Then
        PersonnelDataTables = Array(PERSON_UPN_TABLE, PERSON_APP_TABLE, PERSON_TABLE)
    Else
        PersonnelDataTables = Array(PERSON_APP_TABLE, PERSON_TABLE)
    End If
    db.Close
End Function

' The linked Personnel file opened by itself, for kit_person_upn, which is
' never linked into a front end. Close it when done.
Private Function UpnDb() As DAO.Database
    Dim path As String
    path = PersonnelPath()
    If Len(path) = 0 Then Err.Raise 5, "modPersonnel", "No Personnel file is linked."
    Set UpnDb = DBEngine.OpenDatabase(path, False, False, ";PWD=" & PersonnelKeyFor(path))
End Function

' Make kit_person_upn in udb when it isn't there yet. A new table, so this
' works while other people have the file open (unlike changing kit_person).
Private Sub EnsureUpnTable(udb As DAO.Database)
    If HasTable(udb, PERSON_UPN_TABLE) Then Exit Sub
    CreateUpnTable udb
    udb.TableDefs.Refresh
End Sub

Private Sub CreateUpnTable(db As DAO.Database)
    db.Execute "CREATE TABLE " & PERSON_UPN_TABLE & " (person_id LONG NOT NULL CONSTRAINT pk_" & PERSON_UPN_TABLE & _
        " PRIMARY KEY, upn TEXT(100) NOT NULL, updated_by TEXT(150), updated_at DATETIME)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX ux_person_upn ON " & PERSON_UPN_TABLE & " (upn)", dbFailOnError
End Sub

' What an admin typed as a UPN, trimmed, with a DoD ID made into its UPN.
Private Function UpnFromInput(ByVal s As String) As String
    s = Trim$(s)
    If Len(PersonnelUpnFor(s)) > 0 Then s = PersonnelUpnFor(s)
    UpnFromInput = s
End Function

' Why upn isn't a user principal name ("" when it looks like one).
Private Function UpnProblem(ByVal upn As String) As String
    If Len(upn) > 100 Or InStr(upn, " ") > 0 Or Not upn Like "?*@?*.?*" Then
        UpnProblem = "Enter a user principal name, such as 1234567890" & PERSONNEL_UPN_SUFFIX & ", or a 10-digit DoD ID."
    End If
End Function

' The UPN saved for personId in udb, or "".
Private Function UpnOf(udb As DAO.Database, ByVal personId As Long) As String
    Dim rs As DAO.Recordset
    If Not HasTable(udb, PERSON_UPN_TABLE) Then Exit Function
    Set rs = udb.OpenRecordset("SELECT upn FROM " & PERSON_UPN_TABLE & " WHERE person_id=" & personId, dbOpenSnapshot)
    If Not rs.EOF Then UpnOf = Nz(rs!upn, "")
    rs.Close
End Function

' Whose saved UPN upn is in udb (case ignored), or 0.
Private Function UpnOwner(udb As DAO.Database, ByVal upn As String) As Long
    Dim rs As DAO.Recordset
    If Len(upn) = 0 Then Exit Function
    If Not HasTable(udb, PERSON_UPN_TABLE) Then Exit Function
    Set rs = udb.OpenRecordset("SELECT person_id FROM " & PERSON_UPN_TABLE & " WHERE upn=" & SqlText(upn), dbOpenSnapshot)
    If Not rs.EOF Then UpnOwner = rs!person_id
    rs.Close
End Function

' Save upn as personId's UPN in udb, replacing any they had; "" removes it.
' Makes the table when it isn't there yet. One statement either way, so a
' UPN someone else saved meanwhile fails the unique index and leaves this
' person's old UPN as it was.
Private Sub SaveUpn(udb As DAO.Database, ByVal personId As Long, ByVal upn As String)
    If Len(upn) > 0 Then EnsureUpnTable udb
    If Not HasTable(udb, PERSON_UPN_TABLE) Then Exit Sub
    If Len(upn) = 0 Then
        udb.Execute "DELETE FROM " & PERSON_UPN_TABLE & " WHERE person_id=" & personId, dbFailOnError
    Else
        udb.Execute "UPDATE " & PERSON_UPN_TABLE & " SET upn=" & SqlText(upn) & ", updated_by=" & SqlText(KitUser()) & _
            ", updated_at=Now() WHERE person_id=" & personId, dbFailOnError
        If udb.RecordsAffected = 0 Then
            udb.Execute "INSERT INTO " & PERSON_UPN_TABLE & " (person_id, upn, updated_by, updated_at) VALUES (" & personId & _
                ", " & SqlText(upn) & ", " & SqlText(KitUser()) & ", Now())", dbFailOnError
        End If
    End If
End Sub

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

Private Sub StartSession(ByVal personId As Long, ByVal quiet As Boolean)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT username, display_name, is_admin FROM " & PERSON_TABLE & _
        " WHERE person_id=" & personId, dbOpenSnapshot)
    TempVars("PersonID") = personId
    TempVars("Username") = rs!username.Value
    TempVars("PersonName") = Nz(rs!display_name, rs!username)
    TempVars("PersonIsAdmin") = CBool(rs!is_admin)
    rs.Close
    If Not quiet Then KitToast "Signed in as " & TempVars("PersonName"), "success"
End Sub

' A stored password_hash without the temporary mark.
Private Function HashOnly(ByVal stored As String) As String
    If Left$(stored, 1) = TEMP_PWD_MARK Then HashOnly = Mid$(stored, 2) Else HashOnly = stored
End Function

' A random TEMP_PWD_LEN-character password from TEMP_PWD_CHARS, with at
' least one capital, small letter and digit.
Private Function TemporaryPassword() As String
    Dim hx As String, i As Long, s As String, n As Long
    n = Len(TEMP_PWD_CHARS)
    Do
        hx = RandomHex(TEMP_PWD_LEN)
        s = ""
        For i = 0 To TEMP_PWD_LEN - 1
            s = s & Mid$(TEMP_PWD_CHARS, (CLng("&H" & Mid$(hx, i * 2 + 1, 2)) Mod n) + 1, 1)
        Next
    ' Binary compares: Option Compare Database makes Like and = ignore case.
    Loop Until StrComp(s, LCase$(s), vbBinaryCompare) <> 0 And StrComp(s, UCase$(s), vbBinaryCompare) <> 0 And s Like "*#*"
    TemporaryPassword = s
End Function

' The sign-in name for a Windows account: "NAVY\jsmith" gives "jsmith".
Private Function SignInNameFor(ByVal windowsAccount As String) As String
    SignInNameFor = LCase$(Mid$(windowsAccount, InStrRev(windowsAccount, "\") + 1))
End Function

Private Sub CreatePersonnelTables(db As DAO.Database)
    db.Execute "CREATE TABLE " & PERSON_TABLE & " (person_id COUNTER CONSTRAINT pk_" & PERSON_TABLE & " PRIMARY KEY, " & _
        "windows_account TEXT(150), windows_sid TEXT(184), username TEXT(64) NOT NULL, rank TEXT(20), " & _
        "first_name TEXT(60) NOT NULL, middle_initial TEXT(1), last_name TEXT(60) NOT NULL, display_name TEXT(100), " & _
        "command TEXT(100), email TEXT(150), password_hash TEXT(255), is_admin YESNO, is_active YESNO, " & _
        "created_by TEXT(150), created_at DATETIME, updated_by TEXT(150), updated_at DATETIME)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX ux_person_username ON " & PERSON_TABLE & " (username)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX ux_person_account ON " & PERSON_TABLE & " (windows_account) WITH IGNORE NULL", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX ux_person_sid ON " & PERSON_TABLE & " (windows_sid) WITH IGNORE NULL", dbFailOnError
    db.Execute "CREATE TABLE " & PERSON_APP_TABLE & " (person_id LONG NOT NULL, app_code TEXT(20) NOT NULL, is_active YESNO, " & _
        "granted_at DATETIME, granted_by TEXT(150), CONSTRAINT pk_" & PERSON_APP_TABLE & " PRIMARY KEY (person_id, app_code), " & _
        "CONSTRAINT fk_person_app FOREIGN KEY (person_id) REFERENCES " & PERSON_TABLE & " (person_id))", dbFailOnError
    CreateUpnTable db
End Sub

Private Function HasTable(db As DAO.Database, ByVal nm As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If StrComp(td.Name, nm, vbTextCompare) = 0 Then HasTable = True: Exit Function
    Next
End Function

' Text as a quoted SQL literal.
Private Function SqlText(ByVal s As String) As String
    SqlText = "'" & Replace(s, "'", "''") & "'"
End Function

' Text as a quoted SQL literal, or Null when empty.
Private Function SqlOrNull(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then SqlOrNull = "Null" Else SqlOrNull = SqlText(Trim$(s))
End Function
