Option Compare Database
Option Explicit

' ANCHOR (Access edition) - routing and release: the chains a draft follows
' (routing workflows and their steps), who may release (release
' authorities), the command's transmission settings, and the people who hold
' each step, from the shared Personnel file. The Routing page (Admin:
' Routing, /admin/routing) edits them; the Drafting page reads them
' (modDrafting). The rules are the web app's lib/routing.ts:
'
'   - A draft takes the most specific active workflow that matches its
'     message type and originator (WorkflowFor): a workflow naming types
'     scores 2 when the type is one of them, one naming originators scores 1
'     when the originator is one of them, and a workflow naming types or
'     originators the draft doesn't have doesn't match. Ties go to the oldest.
'   - Every step is held by a person in Personnel who may open ANCHOR, and
'     the last step, and only the last, is the release step
'     (WorkflowProblem). A step whose holder can't act any more (turned off
'     in Personnel, no ANCHOR access) strands the chain until an
'     administrator names the relief (StrandedText).
'   - Only the person holding the release step, who is a release authority
'     for the message type (tblReleaseAuthorities, Scope blank for every
'     type) and holds the Designated Release Authority role, may release it
'     (ReleaseGateProblem).
'   - One row of tblCommandSettings governs every message the command
'     transmits: the text case, the FM line, the MSGID originator and the
'     sequence (prefix, digits, next number), assigned at release in release
'     order (modDrafting.ReleaseDraft).
'
' The web app's "Act as" prototype control has no place here: the signed-in
' person acts as themselves.
'
' Every change on the Routing page checks the System Admin right in the data
' file at the moment of the write (RequireAnchorAdmin) and is audited.
' Expected refusals (a duplicate name, a chain with no release step) come
' back as the sentence to show; faults raise.
'
' Owner: drafting and release. Depends on: modAudit, modErrors, modGlobals,
' modSignIn; the kit's modKitLocks (KitExecute), modPersonnel.

Private Const SETTINGS_ROW As Long = 1
Private Const DEFAULT_ORIGINATOR As String = "COMSUBLANT NORFOLK VA"

' ---------------------------------------------------------------------------
' People (the shared Personnel file)
' ---------------------------------------------------------------------------

' "LT J. Taylor" for a Personnel person_id, "" when there is none (no data
' file, 0 or Null, or not in Personnel any more).
'   pid   a person_id, or Null
' Errors: the engine's.
Public Function RosterName(ByVal pid As Variant) As String
    Dim rs As DAO.Recordset
    If IsNull(pid) Then Exit Function
    If CLng(pid) = 0 Or Not DataLinked() Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT rank, first_name, last_name FROM " & PERSON_TABLE & " WHERE person_id=" & CLng(pid), _
        dbOpenSnapshot)
    If Not rs.EOF Then RosterName = PersonDisplayName(Nz(rs!rank, ""), Nz(rs!first_name, ""), Nz(rs!last_name, ""))
    rs.Close
End Function

' Row source for a person picker: everyone active in Personnel who may open
' ANCHOR, by last name. Columns: person_id, "Taylor, Jordan (LT)".
Public Function RosterSql() As String
    RosterSql = "SELECT p.person_id, p.last_name & ', ' & p.first_name & IIf(Nz(p.rank,'')='','',' (' & p.rank & ')') AS person_name " & _
        "FROM " & PERSON_TABLE & " AS p INNER JOIN " & PERSON_APP_TABLE & " AS a ON p.person_id=a.person_id " & _
        "WHERE p.is_active AND a.is_active AND a.app_code=" & Q(APP_CODE) & " ORDER BY p.last_name, p.first_name"
End Function

' SQL expression: the display name of Personnel row alias ("LT J. Taylor"),
' or what to show when the step has nobody or the person has gone.
'   alias   the kit_person alias in the query
'   idCol   the column holding the person_id the row names
Private Function PersonNameSql(ByVal alias As String, ByVal idCol As String) As String
    PersonNameSql = "IIf(" & idCol & " Is Null,'(nobody yet)',IIf(" & alias & ".person_id Is Null,'(not in Personnel)'," & _
        "IIf(Nz(" & alias & ".rank,'')='',''," & alias & ".rank & ' ') & IIf(Nz(" & alias & ".first_name,'')='',''," & _
        "Left(" & alias & ".first_name,1) & '. ') & " & alias & ".last_name))"
End Function

' Why the person holding a step or a designation can't act on it now, or ""
' when they can: nobody holds it, they are no longer in Personnel, Personnel
' turned them off or took away ANCHOR, or ANCHOR's own access is off.
'   pid        the holder's person_id, or Null or 0 for nobody
'   shownName  the name recorded with the step, for one who has gone
' Errors: the engine's.
Public Function HolderProblem(ByVal pid As Variant, Optional ByVal shownName As String = "") As String
    Dim nm As String
    If Not DataLinked() Then Exit Function
    If Nz(pid, 0) = 0 Then
        HolderProblem = "Nobody holds this step yet. An administrator picks the person under Admin: Routing."
        Exit Function
    End If
    nm = RosterName(pid)
    If Len(nm) = 0 Then
        HolderProblem = IIf(Len(shownName) > 0, shownName, "The person who held it") & _
            " is no longer in the Personnel file. An administrator must name their successor."
    ElseIf Not PersonCanUse(CLng(pid), APP_CODE) Then
        HolderProblem = nm & " is detached from the command (turned off in Personnel, or no longer allowed to open ANCHOR). " & _
            "An administrator must name their relief."
    ElseIf DCount("*", "tblPeople", "PersonID=" & CLng(pid) & " AND Not IsActive") > 0 Then
        HolderProblem = nm & "'s ANCHOR access is turned off. An administrator must name their relief, or turn it back on."
    End If
End Function

' "Held by LT J. Taylor", or HolderProblem's sentence when they can't act.
Public Function HolderNote(ByVal pid As Variant, Optional ByVal shownName As String = "") As String
    HolderNote = HolderProblem(pid, shownName)
    If Len(HolderNote) = 0 Then HolderNote = "Held by " & RosterName(pid)
End Function

' ---------------------------------------------------------------------------
' Lists ("A; B" or "A, B")
' ---------------------------------------------------------------------------

' True when item is one of list's entries (separated by ; or ,), ignoring
' case and spaces around them.
Public Function RoutingListHas(ByVal list As Variant, ByVal item As String) As Boolean
    Dim v As Variant
    item = UCase$(Trim$(item))
    If Len(item) = 0 Then Exit Function
    For Each v In Split(Replace(Nz(list, ""), ",", ";"), ";")
        If UCase$(Trim$(v)) = item Then RoutingListHas = True: Exit Function
    Next
End Function

' A typed list as stored: its entries trimmed, in capitals, blanks and
' repeats left out, joined with "; ". Entries may be separated by ;, a comma
' or a line break.
Public Function RoutingListText(ByVal typed As String) As String
    Dim v As Variant, s As String, e As String
    typed = Replace(Replace(Replace(typed, vbCrLf, ";"), vbLf, ";"), vbCr, ";")
    For Each v In Split(Replace(typed, ",", ";"), ";")
        e = UCase$(Trim$(v))
        If Len(e) > 0 Then
            If Not RoutingListHas(s, e) Then s = s & IIf(Len(s) > 0, "; ", "") & e
        End If
    Next
    RoutingListText = s
End Function

' ---------------------------------------------------------------------------
' Workflows: which applies, and whether it can carry a draft
' ---------------------------------------------------------------------------

' The most specific active workflow for a draft (see the top), 0 for none.
'   msgType     the draft's message type (tblMessageTypes TypeCode)
'   originator  the draft's FM line
' Errors: the engine's.
Public Function WorkflowFor(ByVal msgType As String, ByVal originator As String) As Long
    Dim rs As DAO.Recordset, score As Long, best As Long
    If Not DataLinked() Then Exit Function
    best = -1
    Set rs = CurrentDb.OpenRecordset("SELECT WorkflowID, MatchTypes, MatchOriginators FROM tblWorkflows WHERE IsActive ORDER BY WorkflowID", _
        dbOpenSnapshot)
    Do While Not rs.EOF
        score = 0
        If Len(Trim$(Nz(rs!MatchTypes, ""))) > 0 Then
            If RoutingListHas(rs!MatchTypes, msgType) Then score = score + 2 Else score = -1
        End If
        If score >= 0 And Len(Trim$(Nz(rs!MatchOriginators, ""))) > 0 Then
            If RoutingListHas(rs!MatchOriginators, originator) Then score = score + 1 Else score = -1
        End If
        If score > best Then best = score: WorkflowFor = rs!WorkflowID
        rs.MoveNext
    Loop
    rs.Close
End Function

' The workflow's name, "" when there is none.
Public Function WorkflowName(ByVal wfId As Long) As String
    WorkflowName = Nz(DLookup("WorkflowName", "tblWorkflows", "WorkflowID=" & wfId), "")
End Function

' Why the workflow's chain can't carry a draft, "" when it can: it has no
' steps, a step has no name, or the release step isn't its last and only one.
' Who holds each step is StrandedText's question.
' Errors: the engine's.
Public Function WorkflowProblem(ByVal wfId As Long) As String
    Dim rs As DAO.Recordset, n As Long, nRelease As Long, lastIsRelease As Boolean
    Set rs = CurrentDb.OpenRecordset("SELECT StepTitle, StepAction FROM tblWorkflowSteps WHERE WorkflowID=" & wfId & _
        " ORDER BY StepOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        n = n + 1
        If Len(Trim$(Nz(rs!StepTitle, ""))) = 0 Then WorkflowProblem = "Step " & n & " has no name."
        lastIsRelease = (rs!StepAction = "release")
        If lastIsRelease Then nRelease = nRelease + 1
        rs.MoveNext
    Loop
    rs.Close
    If Len(WorkflowProblem) > 0 Then Exit Function
    If n = 0 Then
        WorkflowProblem = "This chain has no steps."
    ElseIf nRelease = 0 Then
        WorkflowProblem = "This chain has no release step, so nothing would reach a release authority."
    ElseIf nRelease > 1 Or Not lastIsRelease Then
        WorkflowProblem = "The release step must be the last step of the chain, and the only one."
    End If
End Function

' The first step of the workflow that nobody can act on now, as
' "N3 review: <why>", or "" when every step is held by someone who can.
' Errors: the engine's.
Public Function StrandedText(ByVal wfId As Long) As String
    Dim rs As DAO.Recordset, why As String
    Set rs = CurrentDb.OpenRecordset("SELECT StepTitle, PersonID FROM tblWorkflowSteps WHERE WorkflowID=" & wfId & _
        " ORDER BY StepOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        why = HolderProblem(rs!PersonID)
        If Len(why) > 0 Then
            StrandedText = Nz(rs!StepTitle, "A step") & ": " & why
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

' The chain as one line: "N3 review > Chief of Staff review > Commander
' release (release)", with an arrow between the steps.
Public Function WorkflowChainText(ByVal wfId As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT StepTitle, StepAction FROM tblWorkflowSteps WHERE WorkflowID=" & wfId & _
        " ORDER BY StepOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, " " & ChrW(8594) & " ", "") & Nz(rs!StepTitle, "") & IIf(rs!StepAction = "release", " (release)", "")
        rs.MoveNext
    Loop
    rs.Close
    WorkflowChainText = s
End Function

' Rows of the Routing page's workflow list: id, WorkflowName, what it
' matches, and Active or Off.
Public Function WorkflowListSql() As String
    WorkflowListSql = "SELECT WorkflowID AS id, WorkflowName, IIf(Nz(MatchTypes,'')='','Any type',MatchTypes) & ' / ' & " & _
        "IIf(Nz(MatchOriginators,'')='','any originator',MatchOriginators) AS matches_txt, IIf(IsActive,'Active','Off') AS state, " & _
        "Description FROM tblWorkflows ORDER BY WorkflowName"
End Function

' Rows of a workflow's chain: id, n, StepTitle, Review or Release, and who
' holds it.
Public Function WorkflowStepsSql(ByVal wfId As Long) As String
    WorkflowStepsSql = "SELECT s.WfStepID AS id, s.StepOrder AS n, s.StepTitle, IIf(s.StepAction='release','Release','Review') AS action_txt, " & _
        PersonNameSql("p", "s.PersonID") & " AS holder, s.PersonID FROM tblWorkflowSteps AS s LEFT JOIN " & PERSON_TABLE & _
        " AS p ON s.PersonID=p.person_id WHERE s.WorkflowID=" & wfId & " ORDER BY s.StepOrder"
End Function

' ---------------------------------------------------------------------------
' Workflows: changes (Admin: Routing)
' ---------------------------------------------------------------------------

' Add a workflow, off, with a review step and a release step nobody holds
' yet, as the web app's New workflow starts. Returns its id.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function NewWorkflow() As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, nm As String, n As Long, id As Long
    On Error GoTo Fail
    RequireAnchorAdmin
    nm = "New workflow"
    Do While DCount("*", "tblWorkflows", "WorkflowName=" & Q(nm)) > 0
        n = n + 1: nm = "New workflow " & (n + 1)
    Loop
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    id = InsertRow("INSERT INTO tblWorkflows (WorkflowName, Description, IsActive, MatchTypes, MatchOriginators, UpdatedOn, UpdatedBy) " & _
        "VALUES (" & Q(nm) & ", '', False, '', '', Now(), " & Q(Left$(SignedInName(), 100)) & ")")
    KitExecute "INSERT INTO tblWorkflowSteps (WorkflowID, StepOrder, StepTitle, StepAction) VALUES (" & id & ", 1, 'Review', 'review')"
    KitExecute "INSERT INTO tblWorkflowSteps (WorkflowID, StepOrder, StepTitle, StepAction) VALUES (" & id & ", 2, 'Release', 'release')"
    AuditLog "Add workflow", "Routing workflow (" & nm & ")", "Added, turned off until its steps are held."
    ws.CommitTrans: inTrans = False
    NewWorkflow = id
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modRouting.NewWorkflow"
End Function

' Save a workflow's name, description and what it matches. Returns "" when
' saved, else why not (to show).
'   types        the message types it matches, "" for any (RoutingListText)
'   originators  the originators it matches, "" for any
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function SaveWorkflow(ByVal wfId As Long, ByVal nm As String, ByVal desc As String, ByVal types As String, _
        ByVal originators As String) As String
    Dim t As Variant
    nm = Trim$(nm): desc = Trim$(desc)
    types = RoutingListText(types): originators = RoutingListText(originators)
    If Len(nm) = 0 Then SaveWorkflow = "Give the workflow a name.": Exit Function
    If Len(nm) > 100 Then SaveWorkflow = "The name is " & Len(nm) & " characters; keep it to 100.": Exit Function
    If Len(desc) > 255 Then SaveWorkflow = "The description is " & Len(desc) & " characters; keep it to 255.": Exit Function
    If Len(types) > 255 Or Len(originators) > 255 Then SaveWorkflow = "Name fewer message types or originators (255 characters at most).": Exit Function
    If DCount("*", "tblWorkflows", "WorkflowName=" & Q(nm) & " AND WorkflowID<>" & wfId) > 0 Then
        SaveWorkflow = "Another workflow is already called " & nm & ".": Exit Function
    End If
    For Each t In Split(types, "; ")
        If Len(t) > 0 Then
            If DCount("*", "tblMessageTypes", "TypeCode=" & Q(t)) = 0 Then
                SaveWorkflow = t & " isn't in the message type list. Name types as the list has them, such as SAFETY MSG.": Exit Function
            End If
        End If
    Next
    RequireAnchorAdmin
    If KitExecute("UPDATE tblWorkflows SET WorkflowName=" & Q(nm) & ", Description=" & Q(desc) & ", MatchTypes=" & Q(types) & _
            ", MatchOriginators=" & Q(originators) & ", UpdatedOn=Now(), UpdatedBy=" & Q(Left$(SignedInName(), 100)) & _
            " WHERE WorkflowID=" & wfId) = 0 Then
        SaveWorkflow = "This workflow is no longer in the data file.": Exit Function
    End If
    AuditTry "Change workflow", "Routing workflow (" & nm & ")", "Matches " & IIf(Len(types) > 0, types, "any type") & " from " & _
        IIf(Len(originators) > 0, originators, "any originator") & "."
End Function

' Turn a workflow on or off. A workflow whose chain can't carry a draft
' (WorkflowProblem) can't be turned on. Returns "" or why not.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function SetWorkflowActive(ByVal wfId As Long, ByVal isOn As Boolean) As String
    Dim why As String
    If isOn Then
        why = WorkflowProblem(wfId)
        If Len(why) > 0 Then SetWorkflowActive = why & " Fix the chain, then turn it on.": Exit Function
    End If
    RequireAnchorAdmin
    KitExecute "UPDATE tblWorkflows SET IsActive=" & IIf(isOn, "True", "False") & ", UpdatedOn=Now(), UpdatedBy=" & _
        Q(Left$(SignedInName(), 100)) & " WHERE WorkflowID=" & wfId
    AuditTry IIf(isOn, "Turn on workflow", "Turn off workflow"), "Routing workflow (" & WorkflowName(wfId) & ")", _
        IIf(isOn, "Drafts that match it now take it.", "Drafts no longer take it; drafts already in routing keep their route.")
End Function

' Add a step at the end of the chain, or change step stepId (0: add).
' Returns "" or why not. An active workflow keeps a chain that can carry a
' draft: a change that would break it is undone and refused.
'   title   the step's name, "N3 review"
'   action  "review" or "release"
'   pid     who holds it (Personnel person_id), 0 for nobody yet
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function SaveWorkflowStep(ByVal wfId As Long, ByVal stepId As Long, ByVal title As String, ByVal action As String, _
        ByVal pid As Long) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, n As Long
    title = Trim$(title)
    If Len(title) = 0 Then SaveWorkflowStep = "Give the step a name, for example N3 review.": Exit Function
    If Len(title) > 100 Then SaveWorkflowStep = "The step name is " & Len(title) & " characters; keep it to 100.": Exit Function
    If action <> "review" And action <> "release" Then SaveWorkflowStep = "Pick whether the step reviews or releases.": Exit Function
    If pid <> 0 Then
        why = HolderProblem(pid)
        If Len(why) > 0 Then SaveWorkflowStep = why: Exit Function
    End If
    On Error GoTo Fail
    RequireAnchorAdmin
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If stepId = 0 Then
        n = Nz(DMax("StepOrder", "tblWorkflowSteps", "WorkflowID=" & wfId), 0) + 1
        KitExecute "INSERT INTO tblWorkflowSteps (WorkflowID, StepOrder, StepTitle, StepAction, PersonID) VALUES (" & wfId & ", " & n & _
            ", " & Q(title) & ", " & Q(action) & ", " & IIf(pid = 0, "Null", pid) & ")"
    Else
        KitExecute "UPDATE tblWorkflowSteps SET StepTitle=" & Q(title) & ", StepAction=" & Q(action) & ", PersonID=" & _
            IIf(pid = 0, "Null", pid) & " WHERE WfStepID=" & stepId & " AND WorkflowID=" & wfId
    End If
    why = ActiveChainProblem(wfId)
    If Len(why) > 0 Then
        ws.Rollback: inTrans = False
        SaveWorkflowStep = why
        Exit Function
    End If
    AuditLog IIf(stepId = 0, "Add step", "Change step"), "Routing workflow (" & WorkflowName(wfId) & ")", _
        title & " (" & action & "), held by " & IIf(pid = 0, "nobody yet", RosterName(pid)) & "."
    KitExecute "UPDATE tblWorkflows SET UpdatedOn=Now(), UpdatedBy=" & Q(Left$(SignedInName(), 100)) & " WHERE WorkflowID=" & wfId
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modRouting.SaveWorkflowStep", wfId & " " & stepId
End Function

' Take a step off the chain (the others close up). Returns "" or why not.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function RemoveWorkflowStep(ByVal wfId As Long, ByVal stepId As Long) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, title As String
    If DCount("*", "tblWorkflowSteps", "WorkflowID=" & wfId) <= 1 Then
        RemoveWorkflowStep = "A chain needs at least one step: change this one instead.": Exit Function
    End If
    title = Nz(DLookup("StepTitle", "tblWorkflowSteps", "WfStepID=" & stepId), "")
    On Error GoTo Fail
    RequireAnchorAdmin
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "DELETE FROM tblWorkflowSteps WHERE WfStepID=" & stepId & " AND WorkflowID=" & wfId
    RenumberSteps wfId
    why = ActiveChainProblem(wfId)
    If Len(why) > 0 Then
        ws.Rollback: inTrans = False
        RemoveWorkflowStep = why
        Exit Function
    End If
    AuditLog "Remove step", "Routing workflow (" & WorkflowName(wfId) & ")", title & " taken off the chain."
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modRouting.RemoveWorkflowStep", wfId & " " & stepId
End Function

' Move a step one place up (delta -1) or down (+1). Returns "" or why not.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function MoveWorkflowStep(ByVal wfId As Long, ByVal stepId As Long, ByVal delta As Long) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, n As Long, otherId As Long
    n = Nz(DLookup("StepOrder", "tblWorkflowSteps", "WfStepID=" & stepId), 0)
    otherId = Nz(DLookup("WfStepID", "tblWorkflowSteps", "WorkflowID=" & wfId & " AND StepOrder=" & (n + delta)), 0)
    If n = 0 Or otherId = 0 Then Exit Function              ' already first or last
    On Error GoTo Fail
    RequireAnchorAdmin
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "UPDATE tblWorkflowSteps SET StepOrder=" & n & " WHERE WfStepID=" & otherId
    KitExecute "UPDATE tblWorkflowSteps SET StepOrder=" & (n + delta) & " WHERE WfStepID=" & stepId
    why = ActiveChainProblem(wfId)
    If Len(why) > 0 Then
        ws.Rollback: inTrans = False
        MoveWorkflowStep = why
        Exit Function
    End If
    AuditLog "Move step", "Routing workflow (" & WorkflowName(wfId) & ")", Nz(DLookup("StepTitle", "tblWorkflowSteps", "WfStepID=" & stepId), "") & _
        " is now step " & (n + delta) & "."
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modRouting.MoveWorkflowStep", wfId & " " & stepId
End Function

' WorkflowProblem for a workflow that is on (worded for the refusal), ""
' for one that is off: an off workflow may be left half made.
Private Function ActiveChainProblem(ByVal wfId As Long) As String
    Dim why As String
    If DCount("*", "tblWorkflows", "WorkflowID=" & wfId & " AND IsActive") = 0 Then Exit Function
    why = WorkflowProblem(wfId)
    If Len(why) > 0 Then ActiveChainProblem = why & " Turn the workflow off first to rework its chain."
End Function

' Number the workflow's steps 1, 2, 3 in their order.
Private Sub RenumberSteps(ByVal wfId As Long)
    Dim rs As DAO.Recordset, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT StepOrder FROM tblWorkflowSteps WHERE WorkflowID=" & wfId & " ORDER BY StepOrder", dbOpenDynaset)
    Do While Not rs.EOF
        n = n + 1
        If rs!StepOrder <> n Then
            rs.Edit
            rs!StepOrder = n
            rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

' ---------------------------------------------------------------------------
' Release authority
' ---------------------------------------------------------------------------

' The person's release authority designation that covers msgType (Scope
' blank for every type, or naming it), 0 for none.
' Errors: the engine's.
Public Function ReleaseAuthorityFor(ByVal pid As Long, ByVal msgType As String) As Long
    Dim rs As DAO.Recordset
    If pid = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT AuthorityID, Scope FROM tblReleaseAuthorities WHERE PersonID=" & pid & _
        " ORDER BY AuthorityID", dbOpenSnapshot)
    Do While Not rs.EOF
        If Len(Trim$(Nz(rs!Scope, ""))) = 0 Or RoutingListHas(rs!Scope, msgType) Then
            ReleaseAuthorityFor = rs!AuthorityID
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

' Why person pid may not release a message of type msgType now, read from
' the data file, or "" when they may: they need a designation covering the
' type, the Designated Release Authority role, and to be able to act
' (HolderProblem). Holding the release step is the caller's check.
' Errors: the engine's.
Public Function ReleaseGateProblem(ByVal pid As Long, ByVal msgType As String) As String
    Dim why As String
    If ReleaseAuthorityFor(pid, msgType) = 0 Then
        ReleaseGateProblem = "Only the Commander, or a release authority designated for " & msgType & ", may release it."
        Exit Function
    End If
    If Not PersonHasRole(pid, "RELEASE") Then
        ReleaseGateProblem = RosterName(pid) & " is designated, but doesn't hold the " & RoleName("RELEASE") & _
            " role. A System Admin gives it under Admin: Users."
        Exit Function
    End If
    why = HolderProblem(pid)
    If Len(why) > 0 Then ReleaseGateProblem = why & " Designate the relief before releasing."
End Function

' The designation released under ("Chief of Staff"), "" for none.
Public Function ReleaseDesignationOf(ByVal authorityId As Long) As String
    ReleaseDesignationOf = Nz(DLookup("Designation", "tblReleaseAuthorities", "AuthorityID=" & authorityId), "")
End Function

' Rows of the release authority list: id, the person, the designation,
' what they may release, Commander or "", and whether they hold the role.
Public Function AuthorityListSql() As String
    AuthorityListSql = "SELECT a.AuthorityID AS id, " & PersonNameSql("p", "a.PersonID") & " AS holder, a.Designation, " & _
        "IIf(Nz(a.Scope,'')='','All message types',a.Scope) AS scope_txt, IIf(a.IsCommandingOfficer,'Commander','') AS co_txt, " & _
        "IIf(rr.PersonID Is Null,'No role','Holds role') AS role_txt, a.PersonID " & _
        "FROM (tblReleaseAuthorities AS a LEFT JOIN " & PERSON_TABLE & " AS p ON a.PersonID=p.person_id) LEFT JOIN " & _
        "(SELECT pr.PersonID FROM tblPersonRoles AS pr INNER JOIN tblRoles AS r ON pr.RoleID=r.RoleID WHERE r.RoleCode='RELEASE') AS rr " & _
        "ON a.PersonID=rr.PersonID ORDER BY a.IsCommandingOfficer, a.Designation"
End Function

' Designate person pid a release authority. Returns "" or why not.
'   designation  what they release as ("Chief of Staff"); blank takes their
'                billet from their ANCHOR profile
'   isCO         True for the Commander
'   scopeType    the one message type they may release, "" for every type
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function AddReleaseAuthority(ByVal pid As Long, ByVal designation As String, ByVal isCO As Boolean, _
        ByVal scopeType As String) As String
    Dim why As String, crit As String
    designation = Trim$(designation): scopeType = Trim$(scopeType)
    If pid = 0 Then AddReleaseAuthority = "Pick a person from Personnel.": Exit Function
    why = HolderProblem(pid)
    If Len(why) > 0 Then AddReleaseAuthority = why: Exit Function
    If Len(designation) = 0 Then designation = Nz(DLookup("Billet", "tblPeople", "PersonID=" & pid), "")
    If Len(designation) = 0 Then AddReleaseAuthority = "Type what they release as, for example Chief of Staff.": Exit Function
    If Len(designation) > 100 Then AddReleaseAuthority = "The designation is " & Len(designation) & " characters; keep it to 100.": Exit Function
    If Len(scopeType) = 0 Then crit = "(Scope Is Null OR Scope='')" Else crit = "Scope=" & Q(scopeType)
    If DCount("*", "tblReleaseAuthorities", "PersonID=" & pid & " AND " & crit) > 0 Then
        AddReleaseAuthority = RosterName(pid) & " already holds that designation.": Exit Function
    End If
    RequireAnchorAdmin
    KitExecute "INSERT INTO tblReleaseAuthorities (PersonID, Designation, IsCommandingOfficer, Scope, CreatedOn, CreatedBy) VALUES (" & _
        pid & ", " & Q(designation) & ", " & IIf(isCO, "True", "False") & ", " & Q(scopeType) & ", Now(), " & Q(Left$(SignedInName(), 100)) & ")"
    AuditTry "Designate release authority", "Release authority (" & RosterName(pid) & ")", "Releases as " & designation & ", " & _
        IIf(Len(scopeType) > 0, scopeType & " only", "all message types") & "."
End Function

' Take a release authority designation away.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Sub RemoveReleaseAuthority(ByVal authorityId As Long)
    Dim pid As Long, desig As String
    pid = Nz(DLookup("PersonID", "tblReleaseAuthorities", "AuthorityID=" & authorityId), 0)
    desig = ReleaseDesignationOf(authorityId)
    RequireAnchorAdmin
    KitExecute "DELETE FROM tblReleaseAuthorities WHERE AuthorityID=" & authorityId
    AuditTry "Remove release authority", "Release authority (" & RosterName(pid) & ")", "No longer releases as " & desig & "."
End Sub

' ---------------------------------------------------------------------------
' Command transmission settings (tblCommandSettings, one row)
' ---------------------------------------------------------------------------

' One setting: Originator, MsgIdOriginator, TextCase, SeqPrefix, SeqDigits
' or NextSequence; dflt when there is no row (or no data file).
Public Function CommandSetting(ByVal fld As String, ByVal dflt As Variant) As Variant
    If Not DataLinked() Then CommandSetting = dflt: Exit Function
    CommandSetting = Nz(DLookup(fld, "tblCommandSettings", "SettingsID=" & SETTINGS_ROW), dflt)
End Function

' The FM line this command transmits under.
Public Function CommandOriginator() As String
    CommandOriginator = CommandSetting("Originator", DEFAULT_ORIGINATOR)
End Function

' "upper" (every message set in capitals) or "mixed" (as typed).
Public Function CommandTextCase() As String
    CommandTextCase = IIf(CommandSetting("TextCase", "upper") = "mixed", "mixed", "upper")
End Function

' The MSGID originator field (1 to 20 characters, NTP 3 Annex A).
Public Function MsgIdOriginator() As String
    MsgIdOriginator = CommandSetting("MsgIdOriginator", CommandSource())
End Function

' Sequence number n as the serial: "0042" for 4 digits.
Public Function SerialText(ByVal n As Long) As String
    Dim digits As Long
    digits = CLng(CommandSetting("SeqDigits", 4))
    If digits < 1 Or digits > 8 Then digits = 4
    SerialText = Format$(n, String$(digits, "0"))
End Function

' The local sequence for serial: "#0042".
Public Function LocalSequenceText(ByVal serial As String) As String
    LocalSequenceText = CommandSetting("SeqPrefix", "#") & serial
End Function

' The serial before release: "----".
Public Function NullSerial() As String
    Dim digits As Long
    digits = CLng(CommandSetting("SeqDigits", 4))
    If digits < 1 Or digits > 8 Then digits = 4
    NullSerial = String$(digits, "-")
End Function

' The number the next release takes.
Public Function NextSequenceNumber() As Long
    NextSequenceNumber = CLng(CommandSetting("NextSequence", 1))
    If NextSequenceNumber < 1 Then NextSequenceNumber = 1
End Function

' Save the command's transmission settings. Returns "" or why not.
'   textCase    "upper" (every message set in capitals) or "mixed" (as typed)
'   nextTxt     the next sequence number, 1 or more, at most SeqDigits
'               digits; it can't go back to one a released message took
'   originator  the FM line, up to 55 characters (NTP 3 603)
'   msgidOrig   the MSGID originator, 1 to 20 characters (NTP 3 Annex A)
' The prefix and digits stay # and 4: the line after the closing BT is # and
' a four-digit serial (NAVEDTRA 14226), which the naval format check
' requires, so another form would stop every release.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function SaveCommandSettings(ByVal textCase As String, ByVal nextTxt As String, ByVal originator As String, _
        ByVal msgidOrig As String) As String
    Dim used As Long, digits As Long, nxt As Long, sql As String
    originator = UCase$(Trim$(originator)): msgidOrig = UCase$(Trim$(msgidOrig)): nextTxt = Trim$(nextTxt)
    digits = CLng(CommandSetting("SeqDigits", 4))
    If textCase <> "upper" And textCase <> "mixed" Then SaveCommandSettings = "Pick the text case.": Exit Function
    If Len(nextTxt) = 0 Or Len(nextTxt) > digits Or nextTxt Like "*[!0-9]*" Then
        SaveCommandSettings = "The next sequence number is a whole number of at most " & digits & " digits.": Exit Function
    End If
    nxt = CLng(nextTxt)
    If nxt < 1 Then SaveCommandSettings = "The next sequence number is 1 or more.": Exit Function
    used = LastSequenceUsed()
    If nxt <= used Then SaveCommandSettings = "Number " & used & " was already released; the next number must be above it.": Exit Function
    If Len(originator) = 0 Or Len(originator) > 55 Then SaveCommandSettings = "The originator (FM line) is 1 to 55 characters (NTP 3 603).": Exit Function
    If Len(msgidOrig) = 0 Or Len(msgidOrig) > 20 Then SaveCommandSettings = "The MSGID originator is 1 to 20 characters (NTP 3 Annex A).": Exit Function
    RequireAnchorAdmin
    If DCount("*", "tblCommandSettings", "SettingsID=" & SETTINGS_ROW) = 0 Then
        sql = "INSERT INTO tblCommandSettings (SettingsID, Originator, MsgIdOriginator, TextCase, SeqPrefix, SeqDigits, NextSequence, " & _
            "UpdatedOn, UpdatedBy) VALUES (" & SETTINGS_ROW & ", " & Q(originator) & ", " & Q(msgidOrig) & ", " & Q(textCase) & _
            ", '#', 4, " & nxt & ", Now(), " & Q(Left$(SignedInName(), 100)) & ")"
    Else
        sql = "UPDATE tblCommandSettings SET Originator=" & Q(originator) & ", MsgIdOriginator=" & Q(msgidOrig) & ", TextCase=" & _
            Q(textCase) & ", NextSequence=" & nxt & ", UpdatedOn=Now(), UpdatedBy=" & Q(Left$(SignedInName(), 100)) & _
            " WHERE SettingsID=" & SETTINGS_ROW
    End If
    KitExecute sql
    AuditTry "Change settings", "Command transmission settings", "Text " & IIf(textCase = "upper", "in capitals", "as typed") & _
        ", FM " & originator & ", MSGID " & msgidOrig & ", next release " & LocalSequenceText(SerialText(nxt)) & "."
End Function

' The highest sequence number a released draft took, 0 for none.
Private Function LastSequenceUsed() As Long
    Dim rs As DAO.Recordset, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT Msn FROM tblDrafts WHERE Status='released' AND Msn Is Not Null", dbOpenSnapshot)
    Do While Not rs.EOF
        If Len(rs!Msn) > 0 And Len(rs!Msn) <= 9 And Not rs!Msn Like "*[!0-9]*" Then
            n = CLng(rs!Msn)
            If n > LastSequenceUsed Then LastSequenceUsed = n
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function
