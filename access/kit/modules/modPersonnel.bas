Option Compare Database
Option Explicit

' Access UI kit - the shared Personnel directory: who a person is, which apps
' they may open, and signing them in.
'
' The directory is one encrypted data file, Personnel_be.accdb, in its own
' restricted folder on the share, that every app links to. It holds two
' tables, and nothing else holds people's names:
'   kit_person      one row per person: Windows account and SID, sign-in
'                   name, rank and name, command, email, app password hash,
'                   directory admin, active (no SSN or DoD ID)
'   kit_person_app  which apps (app_code "SABER", "SPM", "MG") a person may open
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

' PersonnelStatus results.
Public Const PERSONNEL_READY As Long = 0          ' linked and reachable
Public Const PERSONNEL_NOT_SET_UP As Long = 1     ' no file linked or in the default place: offer find or create
Public Const PERSONNEL_LOST As Long = 2           ' linked, but the file can't be reached: offer find, never create
Public Const PERSONNEL_NO_ADMIN As Long = 3       ' linked and reachable, but nobody in it is an active admin (its
                                                  ' people were deleted, modKitPurge): offer creating the first admin

Private Const DEFAULT_KEY As String = "kit-personnel-1"
Private Const MIN_PASSWORD_LEN As Long = 8

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
    CreatePersonnelFile = AddPerson(firstName, lastName, rank, WinAccount(), "", True)
    KitExecute "UPDATE " & PERSON_TABLE & " SET windows_sid=" & SqlOrNull(WinSid()) & " WHERE person_id=" & CreatePersonnelFile
    SetPersonApp CreatePersonnelFile, appCode, True
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
    RaiseAgain "modPersonnel.CreatePersonnelFile", path
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
    If Not rs.EOF Then ok = VerifyPassword(password, Nz(rs!password_hash, ""))
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
' is refused.
Public Function PersonnelSetPassword(ByVal personId As Long, ByVal password As String) As String
    If Len(password) < MIN_PASSWORD_LEN Then
        PersonnelSetPassword = "Use at least " & MIN_PASSWORD_LEN & " characters."
        Exit Function
    End If
    KitExecute "UPDATE " & PERSON_TABLE & " SET password_hash=" & SqlText(HashPassword(password)) & ", " & StampSql() & _
        " WHERE person_id=" & personId
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

' Add the people in a CSV or Excel file (.csv, .xlsx, .xls; first row is the
' headers) and let them open appCode. Headers in any order and case, with
' spaces and underscores ignored (First name, first_name): First name and
' Last name are needed; Middle initial, Rank, Command, Email, Sign-in name
' (or Username) and Windows account are used when there. A row that can't be
' added (no name, or a sign-in name or Windows account already used) is
' skipped. Returns what happened, for a message.
Public Function PersonnelImport(ByVal path As String, ByVal appCode As String) As String
    Dim rows As Collection, row As Variant, n As Long, added As Long, skipped As Long, why As String, problem As String
    Dim num As Long, d As String, src As String, busy As Boolean
    On Error GoTo Fail
    PersonnelRequireAdmin
    KitBusy "Importing people..."
    busy = True
    Set rows = ImportRows(path)
    For Each row In rows
        n = n + 1
        problem = ImportPerson(row, appCode)
        If Len(problem) = 0 Then
            added = added + 1
        Else
            skipped = skipped + 1
            If skipped <= 5 Then why = why & vbCrLf & "- Row " & (n + 1) & ": " & problem   ' row 1 is the headers
        End If
    Next
    KitBusyEnd
    busy = False
    If n = 0 Then PersonnelImport = "That file has no people in it.": Exit Function
    PersonnelImport = "Added " & added & IIf(added = 1, " person", " people") & "."
    If skipped > 0 Then
        PersonnelImport = PersonnelImport & " Skipped " & skipped & ":" & why
        If skipped > 5 Then PersonnelImport = PersonnelImport & vbCrLf & "- and " & (skipped - 5) & " more"
    End If
    Exit Function
Fail:
    ' KitBusyEnd resets Err, so keep the error and put it back.
    num = Err.Number: d = Err.Description: src = Err.Source
    If busy Then KitBusyEnd
    Err.Number = num: Err.Description = d: Err.Source = src
    RaiseAgain "modPersonnel.PersonnelImport", path
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

' Add one imported row and let them open appCode. "" when added, else why
' not (AddPerson's refusals).
Private Function ImportPerson(ByVal d As Object, ByVal appCode As String) As String
    Dim id As Long, username As String
    On Error GoTo Fail
    username = ImportValue(d, "signinname")
    If Len(username) = 0 Then username = ImportValue(d, "username")
    id = AddPerson(ImportValue(d, "firstname"), ImportValue(d, "lastname"), ImportValue(d, "rank"), _
        ImportValue(d, "windowsaccount"), username, False, ImportValue(d, "middleinitial"), _
        ImportValue(d, "command"), ImportValue(d, "email"))
    SetPersonApp id, appCode, True
    Exit Function
Fail:
    If Err.Number <> 5 Then RaiseAgain "modPersonnel.ImportPerson"
    ImportPerson = Err.Description
End Function

Private Function ImportValue(ByVal d As Object, ByVal key As String) As String
    If d.Exists(key) Then ImportValue = d(key)
End Function

Private Function HeaderKey(ByVal s As String) As String
    HeaderKey = LCase$(Replace(Replace(Replace(Trim$(s), " ", ""), "_", ""), "-", ""))
End Function

' s with Like's wildcards made literal.
Private Function LikeText(ByVal s As String) As String
    LikeText = Replace(Replace(Replace(Replace(s, "[", "[[]"), "*", "[*]"), "?", "[?]"), "#", "[#]")
End Function

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
