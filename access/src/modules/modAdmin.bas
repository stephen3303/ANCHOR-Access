Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Admin page's data: the lists it shows, the
' people with ANCHOR access and their roles (Users), and the message types
' (Message types). Subscription decisions are modSubscriptions'.
'
' People: names, ranks and sign-in names live once in the kit's Personnel
' file (PERSON_TABLE); tblPeople holds each person's ANCHOR profile (billet,
' unit, community, location, DSN), IsActive and IsAdmin, and tblPersonRoles
' their roles. IsAdmin is kept in step with the ADMIN role (System Admin), so
' every admin check reads one column. ANCHOR always keeps one System Admin
' who can sign in (AdmActiveAdminCount): a save that would leave none is
' rolled back.
'
' Message types: tblMessageTypes. A type's code is what messages and
' subscriptions name (tblMessages.MsgType), so it is fixed once added; its
' name, category and description can change, and a type no longer used is
' retired (IsActive off), never deleted, so old traffic keeps its type.
'
' Every write checks the right in the data file at the moment of the write
' (RequireAnchorAdmin) and is audited in the same transaction.
'
' Owner: the Admin page. Depends on: modAudit, modErrors, modGlobals,
' modSignIn; the kit's modKitLocks, modPersonnel.

' The longest values the profile fields take: tblPeople's column sizes
' (backend\modBuildBackEnd.bas).
Public Const ADM_MAX_BILLET As Long = 100
Public Const ADM_MAX_UNIT As Long = 60          ' Unit, Community and Location
Public Const ADM_MAX_DSN As Long = 30
' tblMessageTypes' column sizes.
Public Const ADM_MAX_TYPE_CODE As Long = 30
Public Const ADM_MAX_TYPE_NAME As Long = 100
Public Const ADM_MAX_TYPE_CATEGORY As Long = 30
Public Const ADM_MAX_TYPE_DESC As Long = 255

' ---------------------------------------------------------------------------
' The lists (record sources)
' ---------------------------------------------------------------------------

' SQL for a person's name as the lists show it: "Smith, John (LT)".
'   p   the alias of the Personnel table (PERSON_TABLE) in the query
' Returns an SQL expression. Errors: none.
Public Function AdmPersonNameSql(Optional ByVal p As String = "p") As String
    AdmPersonNameSql = p & ".last_name & ', ' & " & p & ".first_name & IIf(Nz(" & p & ".rank,'')='','',' (' & " & p & ".rank & ')')"
End Function

' The mandatory alert rules (read only), in their order.
Public Function AdmRulesSql() As String
    AdmRulesSql = "SELECT RuleID AS id, RuleKind, MatchText, Audience, 'Within ' & AckHours & ' h' AS ack_txt, Escalate " & _
        "FROM tblAlertRules WHERE IsActive ORDER BY SortOrder, RuleID"
End Function

' The command access groups (read only).
Public Function AdmAccessSql() As String
    AdmAccessSql = "SELECT AccessGroupID AS id, GroupName, Basis, Members, Scope FROM tblAccessGroups ORDER BY GroupName"
End Function

' Everyone with a tblPeople row, by name, with their billet, roles and
' status. withOff False leaves out the people whose access is turned off.
Public Function AdmUsersSql(ByVal withOff As Boolean) As String
    AdmUsersSql = "SELECT e.PersonID AS id, " & AdmPersonNameSql() & " AS person, e.Billet AS billet, " & _
        "AdmPersonRolesText(e.PersonID) AS roles, IIf(e.IsActive,'Active','Deactivated') AS state, p.last_name, p.first_name " & _
        "FROM tblPeople AS e INNER JOIN " & PERSON_TABLE & " AS p ON e.PersonID=p.person_id " & _
        IIf(withOff, "", "WHERE e.IsActive ") & "ORDER BY p.last_name, p.first_name"
End Function

' A person's role names for the users list, "Message Viewer, System Admin",
' or "None". Called once per row by AdmUsersSql.
'   pid   the person's person_id
' Errors: the engine's.
Public Function AdmPersonRolesText(ByVal pid As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT r.RoleName FROM tblPersonRoles AS pr INNER JOIN tblRoles AS r " & _
        "ON pr.RoleID=r.RoleID WHERE pr.PersonID=" & pid & " ORDER BY r.SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ", ", "") & rs!RoleName
        rs.MoveNext
    Loop
    rs.Close
    If Len(s) = 0 Then s = "None"
    AdmPersonRolesText = s
End Function

' The message types by category, then their order.
Public Function AdmTypesSql() As String
    AdmTypesSql = "SELECT TypeID AS id, Category, TypeCode, TypeName, IIf(IsActive,'Active','Retired') AS state " & _
        "FROM tblMessageTypes ORDER BY Category, SortOrder, TypeCode"
End Function

' The categories message types use, for the category box.
Public Function AdmTypeCategoriesSql() As String
    AdmTypeCategoriesSql = "SELECT DISTINCT Category FROM tblMessageTypes WHERE Category Is Not Null ORDER BY Category"
End Function

' The communications classes (read only).
Public Function AdmClassesSql() As String
    AdmClassesSql = "SELECT ClassCode AS id, ClassCode, ClassName, Description FROM tblMessageClasses ORDER BY ClassCode"
End Function

' A person's name as the lists show it, or "person <pid>" when the
' Personnel file no longer has them.
Public Function AdmPersonName(ByVal pid As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT " & AdmPersonNameSql() & " AS nm FROM " & PERSON_TABLE & " AS p WHERE p.person_id=" & pid, _
        dbOpenSnapshot)
    If rs.EOF Then AdmPersonName = "person " & pid Else AdmPersonName = Nz(rs!nm, "person " & pid)
    rs.Close
End Function

' ---------------------------------------------------------------------------
' People
' ---------------------------------------------------------------------------

' System Admins who can sign in now: admin and active in tblPeople, active
' in the Personnel file, with ANCHOR access there. One turned off in the
' Personnel file doesn't count, so the last-admin check keeps one who can
' really use the app.
' Errors: the engine's.
Public Function AdmActiveAdminCount() As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM (tblPeople AS e INNER JOIN " & PERSON_TABLE & " AS p " & _
        "ON e.PersonID=p.person_id) INNER JOIN " & PERSON_APP_TABLE & " AS a ON e.PersonID=a.person_id " & _
        "WHERE e.IsAdmin AND e.IsActive AND p.is_active AND a.is_active AND a.app_code=" & Q(APP_CODE), dbOpenSnapshot)
    AdmActiveAdminCount = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' Give a person from the Personnel file ANCHOR access: a tblPeople row with
' Message Viewer and Subscription Requester (what a first sign-in gives),
' and access to ANCHOR in the Personnel file.
'   pid   the person's person_id
' Returns True when they were added, False when they already had a row.
' Errors: ERR_NOT_ALLOWED unless the signed-in person is a System Admin now;
'   the engine's and KitExecute's (after the rollback).
' Side effects: tblPeople, tblPersonRoles, the Personnel file's app access,
'   and an "Add user" audit row, in one transaction.
Public Function AdmAddPerson(ByVal pid As Long) As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, rs As DAO.Recordset
    RequireAnchorAdmin
    On Error GoTo Fail
    If DCount("*", "tblPeople", "PersonID=" & pid) > 0 Then Exit Function
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "INSERT INTO tblPeople (PersonID, IsAdmin, IsActive, NotifyEmail) VALUES (" & pid & ", False, True, False)"
    Set rs = CurrentDb.OpenRecordset("SELECT RoleID FROM tblRoles WHERE RoleCode IN ('VIEW','REQUEST')", dbOpenSnapshot)
    Do While Not rs.EOF
        KitExecute "INSERT INTO tblPersonRoles (PersonID, RoleID) VALUES (" & pid & ", " & rs!RoleID & ")"
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    If Not PersonCanUse(pid, APP_CODE) Then SetPersonApp pid, APP_CODE, True
    AuditLog "Add user", "User (" & AdmPersonName(pid) & ")", "Given ANCHOR access as Message Viewer and Subscription Requester."
    ws.CommitTrans: inTrans = False
    AdmAddPerson = True
    Exit Function
Fail:
    If Not rs Is Nothing Then rs.Close
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAdmin.AdmAddPerson", pid
End Function

' Save a person's ANCHOR profile, roles and access in one go.
'   pid        the person's person_id (they have a tblPeople row)
'   billet, unit, community, location, dsn   the profile, "" for none
'   roleCodes  the role codes they should hold, ",VIEW,ADMIN," (commas
'              around each); IsAdmin follows ADMIN
'   isActive   whether they may use ANCHOR
' Returns "" when saved, else why it wasn't (nothing is saved then): the
'   save would leave no System Admin who can sign in, or the person may not
'   change that about themselves.
' Errors: ERR_NOT_ALLOWED unless the signed-in person is a System Admin now;
'   the engine's and KitExecute's (after the rollback).
' Side effects: tblPeople and tblPersonRoles, and an audit row for each kind
'   of change ("Change roles", "Deactivate"/"Reactivate", "Update user").
Public Function AdmSavePerson(ByVal pid As Long, ByVal billet As String, ByVal unit As String, ByVal community As String, _
        ByVal location As String, ByVal dsn As String, ByVal roleCodes As String, ByVal isActive As Boolean) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, rs As DAO.Recordset, before As Long
    Dim given As String, taken As String, wasActive As Boolean, wasProfile As String, nowProfile As String, who As String
    Dim isSelf As Boolean, giveAdmin As Boolean
    RequireAnchorAdmin
    On Error GoTo Fail
    isSelf = (pid = SignedInPersonID())
    giveAdmin = InStr(roleCodes, ",ADMIN,") > 0
    If isSelf And Not isActive Then AdmSavePerson = "You can't deactivate your own account. Another System Admin can.": Exit Function
    If isSelf And Not giveAdmin Then AdmSavePerson = "You can't remove your own System Admin role. Another System Admin can.": Exit Function
    who = AdmPersonName(pid)
    DBEngine.Idle dbRefreshCache            ' the counts below read other copies' latest commits
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    ' Admin changes take turns: another admin's save waits here until this
    ' one commits, so the admin count below, read after the wait, sees it.
    If KitExecute("UPDATE tblSettings SET SettingValue=SettingValue WHERE SettingKey='SchemaVersion'") = 0 Then
        LogError "modAdmin.AdmSavePerson", "No SchemaVersion row in tblSettings: admin changes aren't serialised."
    End If
    before = AdmActiveAdminCount()
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblPeople WHERE PersonID=" & pid, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Err.Raise 5, "modAdmin.AdmSavePerson", who & " is no longer on ANCHOR's list of people."
    End If
    wasActive = Nz(rs!IsActive, False)
    wasProfile = ProfileText(Nz(rs!billet, ""), Nz(rs!unit, ""), Nz(rs!community, ""), Nz(rs!location, ""), Nz(rs!dsn, ""))
    rs.Close: Set rs = Nothing
    nowProfile = ProfileText(billet, unit, community, location, dsn)
    KitExecute "UPDATE tblPeople SET Billet=" & NullIfEmpty(billet) & ", Unit=" & NullIfEmpty(unit) & ", Community=" & _
        NullIfEmpty(community) & ", Location=" & NullIfEmpty(location) & ", Dsn=" & NullIfEmpty(dsn) & ", IsActive=" & _
        IIf(isActive, "True", "False") & ", IsAdmin=" & IIf(giveAdmin, "True", "False") & " WHERE PersonID=" & pid
    ' The roles: give the ones ticked that they lack, take the ones unticked.
    Set rs = CurrentDb.OpenRecordset("SELECT r.RoleID, r.RoleCode, r.RoleName, pr.PersonID AS has_it FROM tblRoles AS r LEFT JOIN " & _
        "(SELECT PersonID, RoleID FROM tblPersonRoles WHERE PersonID=" & pid & ") AS pr ON r.RoleID=pr.RoleID ORDER BY r.SortOrder", _
        dbOpenSnapshot)
    Do While Not rs.EOF
        If InStr(roleCodes, "," & rs!RoleCode & ",") > 0 Then
            If IsNull(rs!has_it) Then
                KitExecute "INSERT INTO tblPersonRoles (PersonID, RoleID) VALUES (" & pid & ", " & rs!RoleID & ")"
                given = given & IIf(Len(given) > 0, ", ", "") & rs!RoleName
            End If
        ElseIf Not IsNull(rs!has_it) Then
            KitExecute "DELETE FROM tblPersonRoles WHERE PersonID=" & pid & " AND RoleID=" & rs!RoleID
            taken = taken & IIf(Len(taken) > 0, ", ", "") & rs!RoleName
        End If
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    ' Checked after the write, inside the transaction: two admins taking each
    ' other's admin at once can't both succeed.
    If before > 0 And (Not giveAdmin Or Not isActive) Then
        If AdmActiveAdminCount() = 0 Then
            ws.Rollback: inTrans = False
            AdmSavePerson = who & " is the only System Admin who can sign in, so they keep the System Admin role and their access. " & _
                "Make someone else a System Admin first."
            Exit Function
        End If
    End If
    If Len(given) > 0 Or Len(taken) > 0 Then
        AuditLog "Change roles", "User (" & who & ")", Trim$(IIf(Len(given) > 0, "Given: " & given & ".", "") & _
            IIf(Len(taken) > 0, " Taken away: " & taken & ".", ""))
    End If
    If wasActive <> isActive Then
        AuditLog IIf(isActive, "Reactivate", "Deactivate"), "User (" & who & ")", _
            IIf(isActive, "ANCHOR access turned back on.", "ANCHOR access turned off.")
    End If
    If wasProfile <> nowProfile Then
        AuditLog "Update user", "User (" & who & ")", "Profile was: " & wasProfile & ". Now: " & nowProfile & "."
    End If
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If Not rs Is Nothing Then rs.Close
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAdmin.AdmSavePerson", pid
End Function

' A profile for the audit details: "billet N3A; unit COMSUBLANT N3; ...".
Private Function ProfileText(ByVal billet As String, ByVal unit As String, ByVal community As String, ByVal location As String, _
        ByVal dsn As String) As String
    ProfileText = "billet " & Dash(billet) & "; unit " & Dash(unit) & "; community " & Dash(community) & "; location " & _
        Dash(location) & "; DSN " & Dash(dsn)
End Function

Private Function Dash(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then Dash = "-" Else Dash = Trim$(s)
End Function

' s as an SQL text literal, or Null when it is blank.
Private Function NullIfEmpty(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then NullIfEmpty = "Null" Else NullIfEmpty = Q(Trim$(s))
End Function

' ---------------------------------------------------------------------------
' Message types
' ---------------------------------------------------------------------------

' Why a message type can't be saved as typed, or "".
'   typeId   the type being changed, 0 for a new one (then code is checked too)
Public Function AdmTypeProblem(ByVal typeId As Long, ByVal code As String, ByVal typeTitle As String, ByVal category As String, _
        ByVal descr As String) As String
    If typeId = 0 Then
        If Len(code) = 0 Then AdmTypeProblem = "Type the type's code, as messages name it (GENADMIN).": Exit Function
        If Len(code) > ADM_MAX_TYPE_CODE Then AdmTypeProblem = "The code can be at most " & ADM_MAX_TYPE_CODE & " characters.": Exit Function
        If DCount("*", "tblMessageTypes", "TypeCode=" & Q(code)) > 0 Then
            AdmTypeProblem = "There is already a message type " & code & ".": Exit Function
        End If
    End If
    If Len(typeTitle) = 0 Then AdmTypeProblem = "Type the type's name.": Exit Function
    If Len(typeTitle) > ADM_MAX_TYPE_NAME Then AdmTypeProblem = "The name can be at most " & ADM_MAX_TYPE_NAME & " characters.": Exit Function
    If Len(category) = 0 Then AdmTypeProblem = "Pick or type a category.": Exit Function
    If Len(category) > ADM_MAX_TYPE_CATEGORY Then AdmTypeProblem = "The category can be at most " & ADM_MAX_TYPE_CATEGORY & " characters.": Exit Function
    If Len(descr) > ADM_MAX_TYPE_DESC Then AdmTypeProblem = "The description can be at most " & ADM_MAX_TYPE_DESC & " characters."
End Function

' Add a message type, active, last in its order.
' Returns the new TypeID.
' Errors: ERR_NOT_ALLOWED unless a System Admin now; 5 when AdmTypeProblem
'   finds a problem; the engine's (3022 when another admin added the same
'   code at the same moment).
' Side effects: tblMessageTypes and an "Add message type" audit row, in one
'   transaction.
Public Function AdmAddType(ByVal code As String, ByVal typeTitle As String, ByVal category As String, ByVal descr As String) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, problem As String, id As Long
    RequireAnchorAdmin
    On Error GoTo Fail
    problem = AdmTypeProblem(0, code, typeTitle, category, descr)
    If Len(problem) > 0 Then Err.Raise 5, "modAdmin.AdmAddType", problem
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    id = InsertRow("INSERT INTO tblMessageTypes (TypeCode, TypeName, Category, Description, IsActive, SortOrder) VALUES (" & _
        Q(code) & ", " & Q(typeTitle) & ", " & Q(category) & ", " & NullIfEmpty(descr) & ", True, " & _
        (Nz(DMax("SortOrder", "tblMessageTypes"), 0) + 1) & ")")
    AuditLog "Add message type", "Message type (" & code & ")", typeTitle & ", in " & category & "."
    ws.CommitTrans: inTrans = False
    AdmAddType = id
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAdmin.AdmAddType", code
End Function

' Change a message type's name, category and description (its code stays).
' Returns True when something changed (False: nothing to save).
' Errors: as AdmAddType; 5 when the type is gone.
' Side effects: tblMessageTypes and a "Change message type" audit row.
Public Function AdmSaveType(ByVal typeId As Long, ByVal typeTitle As String, ByVal category As String, ByVal descr As String) As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, problem As String, rs As DAO.Recordset, code As String, was As String, nowText As String
    RequireAnchorAdmin
    On Error GoTo Fail
    problem = AdmTypeProblem(typeId, "", typeTitle, category, descr)
    If Len(problem) > 0 Then Err.Raise 5, "modAdmin.AdmSaveType", problem
    Set rs = CurrentDb.OpenRecordset("SELECT TypeCode, TypeName, Category, Description FROM tblMessageTypes WHERE TypeID=" & typeId, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Err.Raise 5, "modAdmin.AdmSaveType", "This message type is no longer there."
    End If
    code = rs!TypeCode
    was = Nz(rs!TypeName, "") & " | " & Nz(rs!category, "") & " | " & Nz(rs!Description, "")
    rs.Close: Set rs = Nothing
    nowText = typeTitle & " | " & category & " | " & descr
    If StrComp(was, nowText, vbBinaryCompare) = 0 Then Exit Function
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "UPDATE tblMessageTypes SET TypeName=" & Q(typeTitle) & ", Category=" & Q(category) & ", Description=" & _
        NullIfEmpty(descr) & " WHERE TypeID=" & typeId
    AuditLog "Change message type", "Message type (" & code & ")", "Was: " & was & ". Now: " & nowText & "."
    ws.CommitTrans: inTrans = False
    AdmSaveType = True
    Exit Function
Fail:
    If Not rs Is Nothing Then rs.Close
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAdmin.AdmSaveType", typeId
End Function

' Retire a message type (no longer offered for new traffic, filters and
' drafting) or restore it.
'   active   False to retire, True to restore
' Errors: ERR_NOT_ALLOWED unless a System Admin now; the engine's.
' Side effects: tblMessageTypes.IsActive and a "Retire message type" or
'   "Restore message type" audit row, in one transaction.
Public Sub AdmSetTypeActive(ByVal typeId As Long, ByVal active As Boolean)
    Dim ws As DAO.Workspace, inTrans As Boolean, code As String
    RequireAnchorAdmin
    On Error GoTo Fail
    code = Nz(DLookup("TypeCode", "tblMessageTypes", "TypeID=" & typeId), "")
    If Len(code) = 0 Then Err.Raise 5, "modAdmin.AdmSetTypeActive", "This message type is no longer there."
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "UPDATE tblMessageTypes SET IsActive=" & IIf(active, "True", "False") & " WHERE TypeID=" & typeId
    AuditLog IIf(active, "Restore message type", "Retire message type"), "Message type (" & code & ")", _
        IIf(active, "Offered again for new traffic, filters and drafting.", "No longer offered for new traffic; messages of this type keep it.")
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAdmin.AdmSetTypeActive", typeId
End Sub
