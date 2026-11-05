Option Compare Database
Option Explicit

' ANCHOR (Access edition) - start-up, sign-in and roles.
'
' frmMain's Form_Load runs AppStartUp under the splash:
'   1. the shared data file: KitEnsureBackend finds it next to the front end or
'      asks for it when it has moved;
'   2. its version (tblSettings SchemaVersion) against DATA_VERSION, so an old
'      copy stops against a data file a newer copy has changed;
'   3. sign-in: KitPersonnelStart finds or sets up the shared Personnel file and
'      signs the Windows user in (TempVars PersonID, PersonName). The web app's
'      CAC sign-in is the Windows account here;
'   4. the person's ANCHOR roles (LoadRoles) into TempVars ANRoles
'      (",VIEW,REQUEST,BRIEF,") and ANIsAdmin. Everyone the Personnel file lets
'      use ANCHOR is a Message Viewer and Subscription Requester: their first
'      sign-in adds their tblPeople row with those two roles. A person turned
'      off in tblPeople is told why and the app closes;
'   5. the help (KitHelpSync), and the audit rows not sealed yet
'      (modAudit.AuditSealPending: a new data file's seeded rows).
' Then ShowPagesForRoles hides the sidebar pages they can't use, and
' frmMain.RoutePage sends any other way in (Ctrl+K) to Home.
'
' Roles (tblRoles, the web app's names): VIEW Message Viewer, REQUEST
' Subscription Requester, BRIEF Briefing Creator, REPORT Report Viewer,
' APPROVER Subscription Approver, ADMIN System Admin. tblPeople.IsAdmin is
' kept in step with ADMIN, so the kit-style admin checks read one column.
'
' The first Personnel admin to open a new data file (tblPeople still empty)
' becomes its admin with every role, once (tblSettings AdminBootstrapped).
'
' A front end built with no data file (build.ps1 with no -BackEnd, for UI work)
' has no tblSettings: it skips all of this and shows every page.
'
' The TempVars only decide what the sidebar shows. Every write that needs a
' right checks it again in the data file at the moment of the write
' (RequireAnchorAdmin, RequireRole), raising ERR_NOT_ALLOWED, which
' HandleError shows as the plain reason, not as a fault.
'
' Owner: start-up and sign-in. Depends on: modAudit, modErrors, modGlobals;
' the kit's modKitBackend, modKitHelp, modKitLocks, modKitShell, modKitUI,
' modKitUtil, modPersonnel, modPersonnelUI.

Public Const DATA_VERSION As Long = 1            ' the tblSettings SchemaVersion this copy reads

' Raised by the checks made at the point of a write when the signed-in person
' may not make it. Err.Description is the reason, written for the user.
Public Const ERR_NOT_ALLOWED As Long = vbObjectError + 70

Public Const SET_ADMIN_BOOTSTRAPPED As String = "AdminBootstrapped"
Private Const PROBE_TABLE As String = "tblSettings"

' Whether the signed-in person may open a page, from the roles read at
' sign-in (TempVars). It only decides what the sidebar shows and where
' Ctrl+K may go; every write checks the right again in the data file.
'   key   a sidebar key, or a key only frmMain.PageForm knows (message)
' Returns True for every page in a UI-only build (no data file).
' Errors: none.
Public Function PageAllowed(ByVal key As String) As Boolean
    If Not DataLinked() Then PageAllowed = True: Exit Function
    Select Case key
        Case "home", "messages", "message", "search", "subscriptions", "watch", "profile": PageAllowed = True
        Case "briefings": PageAllowed = HasRole("BRIEF") Or IsAnchorAdmin()
        Case "reports": PageAllowed = HasRole("REPORT") Or IsAnchorAdmin()
        Case "admin": PageAllowed = HasRole("APPROVER") Or IsAnchorAdmin()
        Case "audit": PageAllowed = IsAnchorAdmin()
    End Select
End Function

' Start-up, run by frmMain.Form_Load under the splash (see the top).
' Returns True when the app may open, False when it should close.
' Errors: the kit's and the engine's pass up to frmMain's handler.
Public Function AppStartUp() As Boolean
    Dim problem As String
    If Not DataLinked() Then AppStartUp = True: Exit Function
    If Not TableExists(PROBE_TABLE) Then
        KitMsgBox "This copy of " & APP_NAME & " has lost its link to the data file, so it can't open. " & _
            "Get a fresh copy from your administrator.", vbExclamation, APP_NAME
        Exit Function
    End If
    SplashStatus "Opening the data file" & ChrW(8230)
    If Not KitEnsureBackend(PROBE_TABLE, BE_FILE) Then Exit Function
    problem = DataVersionProblem()
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, APP_NAME
        Exit Function
    End If
    If SignedInPersonID() = 0 Then          ' a palette switch reopens the shell already signed in
        If Not KitPersonnelStart(PROBE_TABLE, APP_CODE) Then Exit Function
    End If
    SplashStatus "Loading your roles" & ChrW(8230)
    problem = LoadRoles()
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, APP_NAME
        Exit Function
    End If
    SplashStatus "Loading help" & ChrW(8230)
    SyncHelp
    SealAudit
    AppStartUp = True
End Function

' Start-up step 5. Help only helps: when it can't be read the error is logged.
Private Sub SyncHelp()
    On Error GoTo Fail
    KitHelpSync
    Exit Sub
Fail:
    LogError "modSignIn.SyncHelp"
End Sub

' Start-up step 5: seal audit rows with no hash yet. A failure is logged; the
' Audit page's Verify then reports the rows as unsealed.
Private Sub SealAudit()
    On Error GoTo Fail
    AuditSealPending
    Exit Sub
Fail:
    LogError "modSignIn.SealAudit"
End Sub

' Hide the sidebar pages the signed-in person can't use (PageAllowed).
Public Sub ShowPagesForRoles(frm As Form)
    Dim k As Variant
    For Each k In Array("briefings", "reports", "admin", "audit")
        ShellShowItem frm, CStr(k), PageAllowed(CStr(k))
    Next
End Sub

' Sign out and show the sign-in screen (Profile's Sign out). Only with a
' data file and while the current page can be left. Signing in reopens
' frmMain on Home, whose Form_Load reads the new person's roles.
Public Function AppSignOut()
    Dim v As Variant
    On Error GoTo Err_Handler
    If Not DataLinked() Then Exit Function
    If IsFormOpen("frmMain") Then
        If Not Forms("frmMain").CanLeavePage() Then Exit Function
    End If
    AuditLog "Sign out", "Session", "Signed out of " & APP_NAME & "."
    ClearRoles
    On Error Resume Next                    ' one that was never set
    For Each v In Array("OpenMessageID", "MessagesFilter", "SearchText")
        TempVars.Remove v
    Next
    On Error GoTo Err_Handler
    If Not KitPersonnelSignOut(APP_CODE) Then
        DoCmd.Quit acQuitSaveNone
        Exit Function
    End If
    If IsFormOpen("dlgToast") Then DoCmd.Close acForm, "dlgToast", acSaveNo
    If IsFormOpen("frmMain") Then Forms("frmMain").Reopen "home"
    Exit Function
Err_Handler:
    HandleError "modSignIn.AppSignOut"
End Function

' True when this front end belongs with a data file: built with one, or
' linked to its tables. False only for a UI-only build.
Public Function DataLinked() As Boolean
    DataLinked = BuiltWithDataFile() Or TableExists(PROBE_TABLE)
End Function

Private Function BuiltWithDataFile() As Boolean
    On Error Resume Next                    ' error 3270: no such property
    BuiltWithDataFile = CurrentDb.Properties("HasDataFile")
End Function

' Why this copy can't use its data file (KitVersionProblem), or "".
' Errors: the engine's, when tblSettings can't be read.
Public Function DataVersionProblem() As String
    Dim v As Long
    v = CLng(Nz(DLookup("SettingValue", PROBE_TABLE, "SettingKey='SchemaVersion'"), "0"))
    If v = 0 Then Exit Function
    DataVersionProblem = KitVersionProblem(v, DATA_VERSION, "", "")
End Function

' ---------------------------------------------------------------------------
' Roles
' ---------------------------------------------------------------------------

' Read the signed-in person's ANCHOR access into TempVars ANRoles and
' ANIsAdmin (cleared first), adding their tblPeople row on first sign-in.
' Returns "" when they may use the app, else what to tell them.
' Errors: the engine's.
Public Function LoadRoles() As String
    Dim pid As Long, rs As DAO.Recordset, roles As String
    pid = SignedInPersonID()
    ClearRoles
    If pid = 0 Then LoadRoles = "Nobody is signed in.": Exit Function
    If DCount("*", "tblPeople", "PersonID=" & pid) = 0 Then AddSelf pid
    Set rs = CurrentDb.OpenRecordset("SELECT IsAdmin, IsActive FROM tblPeople WHERE PersonID=" & pid, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        LoadRoles = "You're signed in as " & SignedInName() & ", but ANCHOR couldn't add you to its list of people. Try again."
        Exit Function
    End If
    If Not Nz(rs!IsActive, False) Then
        rs.Close
        LoadRoles = "You're signed in as " & SignedInName() & ", but your " & APP_NAME & " access is turned off. " & _
            "Ask a System Admin to turn it on under Admin: Users."
        Exit Function
    End If
    TempVars("ANIsAdmin") = CBool(Nz(rs!IsAdmin, False))
    rs.Close
    roles = PersonRoleCodes(pid)
    TempVars("ANRoles") = "," & roles & IIf(Len(roles) > 0, ",", "")
    KitExecute "UPDATE tblPeople SET LastSeenOn=Now() WHERE PersonID=" & pid
End Function

' The signed-in person's name for messages ("LT J. Smith"), or "you".
Public Function SignedInName() As String
    On Error Resume Next
    SignedInName = Nz(TempVars("PersonName"), "you")
    If Len(SignedInName) = 0 Then SignedInName = "you"
End Function

' First sign-in: a tblPeople row with Message Viewer and Subscription
' Requester. On a new data file (nobody in tblPeople, no AdminBootstrapped)
' a Personnel admin becomes its admin with every role, once. Another copy
' adding the same person at the same moment (3022) is left to win.
Private Sub AddSelf(ByVal pid As Long)
    Dim ws As DAO.Workspace, inTrans As Boolean, first As Boolean, rs As DAO.Recordset
    On Error GoTo Fail
    first = (DCount("*", "tblPeople") = 0 And SettingText(SET_ADMIN_BOOTSTRAPPED) <> "1")
    If first Then first = PersonnelIsAdminNow()
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "INSERT INTO tblPeople (PersonID, IsAdmin, IsActive, NotifyEmail) VALUES (" & pid & ", " & _
        IIf(first, "True", "False") & ", True, False)"
    Set rs = CurrentDb.OpenRecordset("SELECT RoleID, RoleCode FROM tblRoles", dbOpenSnapshot)
    Do While Not rs.EOF
        If first Or rs!RoleCode = "VIEW" Or rs!RoleCode = "REQUEST" Then
            KitExecute "INSERT INTO tblPersonRoles (PersonID, RoleID) VALUES (" & pid & ", " & rs!RoleID & ")"
        End If
        rs.MoveNext
    Loop
    rs.Close
    If first Then
        If DCount("*", "tblSettings", "SettingKey=" & Q(SET_ADMIN_BOOTSTRAPPED)) = 0 Then
            KitExecute "INSERT INTO tblSettings (SettingKey, SettingValue, Description) VALUES (" & Q(SET_ADMIN_BOOTSTRAPPED) & _
                ", '1', 'Set once the first admin was made from the Personnel admins')"
        Else
            KitExecute "UPDATE tblSettings SET SettingValue='1' WHERE SettingKey=" & Q(SET_ADMIN_BOOTSTRAPPED)
        End If
    End If
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    If Err.Number = 3022 Then Exit Sub
    RaiseAgain "modSignIn.AddSelf", pid
End Sub

' True when the signed-in person held a role at sign-in (TempVars ANRoles).
'   code   VIEW, REQUEST, BRIEF, REPORT, APPROVER or ADMIN
Public Function HasRole(ByVal code As String) As Boolean
    On Error Resume Next                    ' TempVars may be unset
    HasRole = (InStr(Nz(TempVars("ANRoles"), ""), "," & code & ",") > 0)
End Function

' Raise ERR_NOT_ALLOWED unless the signed-in person holds role code now
' (read from the data file), or is an admin. A front end with no data file
' passes.
Public Sub RequireRole(ByVal code As String, ByVal what As String)
    Dim pid As Long
    If Not DataLinked() Then Exit Sub
    pid = SignedInPersonID()
    If pid <> 0 Then
        If DCount("*", "tblPeople", "PersonID=" & pid & " AND IsActive") > 0 Then
            If PersonHasRole(pid, code) Or PersonHasRole(pid, "ADMIN") Then Exit Sub
        End If
    End If
    Err.Raise ERR_NOT_ALLOWED, "modSignIn.RequireRole", "Only someone with the " & RoleName(code) & " role can " & what & _
        ". Your roles may have changed since you signed in."
End Sub

' Raise ERR_NOT_ALLOWED unless the signed-in person is an active ANCHOR
' admin now, read from the data file.
Public Sub RequireAnchorAdmin()
    Dim pid As Long, ok As Boolean
    If Not DataLinked() Then Exit Sub
    pid = SignedInPersonID()
    If pid <> 0 Then
        ok = DCount("*", "tblPeople", "PersonID=" & pid & " AND IsAdmin AND IsActive") > 0
        If ok Then ok = PersonCanUse(pid, APP_CODE)
    End If
    TempVars("ANIsAdmin") = ok
    If Not ok Then Err.Raise ERR_NOT_ALLOWED, "modSignIn.RequireAnchorAdmin", _
        "Only an active System Admin can make this change. Your admin rights may have been taken away since you signed in."
End Sub

' True when the signed-in person was an admin when last checked.
Public Function IsAnchorAdmin() As Boolean
    On Error Resume Next
    IsAnchorAdmin = Nz(TempVars("ANIsAdmin"), False)
End Function

Private Sub ClearRoles()
    Dim v As Variant
    On Error Resume Next                    ' one that was never set
    For Each v In Array("ANRoles", "ANIsAdmin")
        TempVars.Remove v
    Next
End Sub

' ---------------------------------------------------------------------------
' Role and setting look-ups
' ---------------------------------------------------------------------------

Public Function RoleID(ByVal code As String) As Long
    RoleID = Nz(DLookup("RoleID", "tblRoles", "RoleCode=" & Q(code)), 0)
End Function

Public Function RoleName(ByVal code As String) As String
    RoleName = Nz(DLookup("RoleName", "tblRoles", "RoleCode=" & Q(code)), code)
End Function

Public Function PersonHasRole(ByVal pid As Long, ByVal code As String) As Boolean
    PersonHasRole = DCount("*", "tblPersonRoles", "PersonID=" & pid & " AND RoleID=" & RoleID(code)) > 0
End Function

' The role codes a person holds, in tblRoles order, "VIEW,REQUEST", or "".
Public Function PersonRoleCodes(ByVal pid As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT r.RoleCode FROM tblPersonRoles AS pr INNER JOIN tblRoles AS r " & _
        "ON pr.RoleID=r.RoleID WHERE pr.PersonID=" & pid & " ORDER BY r.SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ",", "") & rs!RoleCode
        rs.MoveNext
    Loop
    rs.Close
    PersonRoleCodes = s
End Function

' A tblSettings value as text, dflt when there is no row.
Public Function SettingText(ByVal key As String, Optional ByVal dflt As String = "") As String
    If Not DataLinked() Then SettingText = dflt: Exit Function
    SettingText = Nz(DLookup("SettingValue", "tblSettings", "SettingKey=" & Q(key)), dflt)
End Function

' This command's own source code (COMSUBLANT): its traffic is "sent by us".
Public Function CommandSource() As String
    CommandSource = SettingText("CommandSource", "COMSUBLANT")
End Function
