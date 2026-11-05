Option Compare Database
Option Explicit

' ANCHOR (Access edition) - subscriptions: the catalogue (tblSubscriptions),
' each person's subscriptions and requests (tblPersonSubs), and the steps
' that change them.
'
' A person's status for a subscription is their latest tblPersonSubs row
' for it (highest PersonSubID):
'   Active            they receive it          (Pause)
'   Paused            they paused it           (Resume)
'   Pending Approval  they asked for it        (Withdraw; an approver decides)
'   Denied            an approver said no      (Request again)
'   no row            Not subscribed           (Request)
' A mandatory subscription (IsMandatory) is assigned by command: it always
' shows Active, needs no row, and can't be requested, paused or withdrawn.
'
' Who may do what is checked again in the data file at the moment of the
' write (modSignIn): RequestSubscription needs Subscription Requester
' (REQUEST); ApproveSubscription and DenySubscription need Subscription
' Approver (APPROVER). Pause, Resume and Withdraw change only the signed-in
' person's own rows. Every change is audited (modAudit), in the same
' transaction as the write.
'
' For the Admin page (approvals): ApproveSubscription(personSubId, decisionNote)
' and DenySubscription(personSubId, decisionNote), each returning "" or what to
' tell the approver. PendingSubsSql lists the requests waiting.
'
' Owner: the Subscriptions page. Depends on: modAudit, modErrors,
' modGlobals, modMessages, modSignIn; the kit's modKitLocks (KitExecute),
' modPersonnel.

Public Const SUB_NOT_SUBSCRIBED As String = "Not subscribed"

' ---------------------------------------------------------------------------
' Reading
' ---------------------------------------------------------------------------

' SQL expression: the signed-in person's status of subscription s (alias s
' over tblSubscriptions): Active for a mandatory one, else their latest
' row's Status, else "Not subscribed".
Private Function MyStatusSql() As String
    MyStatusSql = "IIf(s.IsMandatory,'Active',Nz((SELECT TOP 1 ps.Status FROM tblPersonSubs AS ps WHERE ps.SubID=s.SubID AND ps.PersonID=" & _
        SignedInPersonID() & " ORDER BY ps.PersonSubID DESC),'" & SUB_NOT_SUBSCRIBED & "'))"
End Function

' The catalogue for the Subscriptions page: active subscriptions in
' category cat ("" for all) whose name, description or category contain
' find ("" for any), mandatory first, then by name. Columns: id, SubName,
' Category, assign_txt ("Mandatory" or ""), my_status.
Public Function SubsListSql(ByVal cat As String, ByVal find As String) As String
    Dim w As String, p As String
    w = "s.IsActive"
    If Len(cat) > 0 Then w = w & " AND s.Category=" & Q(cat)
    If Len(find) > 0 Then
        p = Q("*" & LikeLiteral(find) & "*")
        w = w & " AND (s.SubName LIKE " & p & " OR s.Description LIKE " & p & " OR s.Category LIKE " & p & ")"
    End If
    SubsListSql = "SELECT s.SubID AS id, s.SubName, s.Category, IIf(s.IsMandatory,'Mandatory','') AS assign_txt, " & _
        MyStatusSql() & " AS my_status FROM tblSubscriptions AS s WHERE " & w & " ORDER BY s.IsMandatory, s.SubName"
End Function

' The active catalogue's categories, by name, as a combo box value list
' with "(All categories)" first. Each item is quoted, so a name with ";" or
' "," stays one item.
Public Function SubCategoriesList() As String
    Dim rs As DAO.Recordset, s As String
    s = """(All categories)"""
    If DataLinked() Then
        Set rs = CurrentDb.OpenRecordset("SELECT DISTINCT Category FROM tblSubscriptions WHERE IsActive AND Category Is Not Null " & _
            "ORDER BY Category", dbOpenSnapshot)
        Do While Not rs.EOF
            s = s & ";""" & Replace(rs!Category, """", """""") & """"
            rs.MoveNext
        Loop
        rs.Close
    End If
    SubCategoriesList = s
End Function

' The signed-in person's latest tblPersonSubs row for subscription subId,
' 0 when they have none.
Public Function MyPersonSubID(ByVal subId As Long) As Long
    Dim pid As Long
    pid = SignedInPersonID()
    If pid = 0 Then Exit Function
    MyPersonSubID = Nz(DMax("PersonSubID", "tblPersonSubs", "SubID=" & subId & " AND PersonID=" & pid), 0)
End Function

' The signed-in person's status of subscription subId: Active, Paused,
' Pending Approval, Denied or "Not subscribed" (SUB_NOT_SUBSCRIBED).
' Mandatory ones are always Active.
Public Function MySubStatus(ByVal subId As Long) As String
    Dim psId As Long
    If Nz(DLookup("IsMandatory", "tblSubscriptions", "SubID=" & subId), False) Then MySubStatus = "Active": Exit Function
    psId = MyPersonSubID(subId)
    If psId = 0 Then MySubStatus = SUB_NOT_SUBSCRIBED: Exit Function
    MySubStatus = Nz(DLookup("Status", "tblPersonSubs", "PersonSubID=" & psId), SUB_NOT_SUBSCRIBED)
End Function

' A "; "-separated list from the catalogue (MsgTypes, Originators,
' Keywords, Precedences) as an array of trimmed items, none blank.
Private Function ListItems(ByVal v As Variant) As Variant
    Dim parts As Variant, i As Long, s As String
    parts = Split(Replace(Nz(v, ""), ",", ";"), ";")
    For i = 0 To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then s = s & vbTab & Trim$(parts(i))
    Next
    If Len(s) = 0 Then ListItems = Split("") Else ListItems = Split(Mid$(s, 2), vbTab)
End Function

' Items as an SQL IN list: ('A','B').
Private Function InList(ByVal items As Variant) As String
    Dim i As Long, s As String
    For i = 0 To UBound(items)
        s = s & "," & Q(items(i))
    Next
    InList = "(" & Mid$(s, 2) & ")"
End Function

' The WHERE condition (over tblMessages AS m) for the messages subscription
' subId matches, as the web app matches them: any of its message types, any
' of its precedences, from any of its originators (the message's source, or
' an originator that starts with it), and with any of its keywords in the
' subject or communities. A criterion left empty matches every message;
' "False" when the subscription is gone.
Public Function SubMatchWhereSql(ByVal subId As Long) As String
    Dim rs As DAO.Recordset, w As String, items As Variant, i As Long, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT MsgTypes, Originators, Keywords, Precedences FROM tblSubscriptions WHERE SubID=" & subId, _
        dbOpenSnapshot)
    If rs.EOF Then rs.Close: SubMatchWhereSql = "False": Exit Function
    items = ListItems(rs!MsgTypes)
    If UBound(items) >= 0 Then w = w & " AND m.MsgType IN " & InList(items)
    items = ListItems(rs!Precedences)
    If UBound(items) >= 0 Then w = w & " AND m.Precedence IN " & InList(items)
    items = ListItems(rs!Originators)
    If UBound(items) >= 0 Then
        s = ""
        For i = 0 To UBound(items)
            s = s & " OR m.Originator LIKE " & Q(LikeLiteral(CStr(items(i))) & "*")
        Next
        w = w & " AND (m.Source IN " & InList(items) & s & ")"
    End If
    items = ListItems(rs!Keywords)
    If UBound(items) >= 0 Then
        s = ""
        For i = 0 To UBound(items)
            s = s & " OR m.Subject LIKE " & Q("*" & LikeLiteral(CStr(items(i))) & "*") & _
                " OR m.Communities LIKE " & Q("*" & LikeLiteral(CStr(items(i))) & "*")
        Next
        w = w & " AND (" & Mid$(s, 5) & ")"
    End If
    rs.Close
    If Len(w) = 0 Then SubMatchWhereSql = "True" Else SubMatchWhereSql = Mid$(w, 6)
End Function

' How many messages in ANCHOR subscription subId matches (SubMatchWhereSql).
Public Function SubMatchCount(ByVal subId As Long) As Long
    Dim rs As DAO.Recordset
    If Not DataLinked() Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM tblMessages AS m WHERE " & SubMatchWhereSql(subId), dbOpenSnapshot)
    SubMatchCount = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' What subscription subId matches, one criterion a line, for the detail
' card: message types (with their names from tblMessageTypes), originators,
' keywords and precedence; "Any" for an empty one.
Public Function SubCriteriaText(ByVal subId As Long) As String
    Dim rs As DAO.Recordset, items As Variant, i As Long, s As String, nm As String
    Set rs = CurrentDb.OpenRecordset("SELECT MsgTypes, Originators, Keywords, Precedences FROM tblSubscriptions WHERE SubID=" & subId, _
        dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    items = ListItems(rs!MsgTypes)
    For i = 0 To UBound(items)
        nm = Nz(DLookup("TypeName", "tblMessageTypes", "TypeCode=" & Q(items(i))), "")
        s = s & IIf(i > 0, ", ", "") & items(i) & IIf(Len(nm) > 0, " (" & nm & ")", "")
    Next
    SubCriteriaText = "Message types: " & IIf(Len(s) > 0, s, "Any") & vbCrLf & _
        "Originators: " & ItemsText(rs!Originators) & vbCrLf & _
        "Keywords: " & ItemsText(rs!Keywords) & vbCrLf & _
        "Precedence: " & ItemsText(rs!Precedences)
    rs.Close
End Function

Private Function ItemsText(ByVal v As Variant) As String
    Dim items As Variant
    items = ListItems(v)
    If UBound(items) < 0 Then ItemsText = "Any" Else ItemsText = Join(items, ", ")
End Function

' The delivery of a tblPersonSubs row in words: "Daily Digest; In app, Email".
Public Function SubDeliveryText(ByVal personSubId As Long) As String
    Dim rs As DAO.Recordset, ch As String
    Set rs = CurrentDb.OpenRecordset("SELECT Frequency, ChInApp, ChEmail, ChTeams FROM tblPersonSubs WHERE PersonSubID=" & personSubId, _
        dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    If Nz(rs!ChInApp, False) Then ch = ch & ", In app"
    If Nz(rs!ChEmail, False) Then ch = ch & ", Email"
    If Nz(rs!ChTeams, False) Then ch = ch & ", Teams"
    SubDeliveryText = Nz(rs!Frequency, "Real-time") & IIf(Len(ch) > 0, "; " & Mid$(ch, 3), "")
    rs.Close
End Function

' The requests waiting for an approver, oldest first, for the Admin page:
' id (PersonSubID), SubName, PersonName, Frequency, Justification,
' RequestedOn and requested_txt.
Public Function PendingSubsSql() As String
    PendingSubsSql = "SELECT ps.PersonSubID AS id, s.SubName, ps.PersonName, ps.Frequency, ps.Justification, ps.RequestedOn, " & _
        "Format(ps.RequestedOn,'dd mmm yyyy') AS requested_txt FROM tblPersonSubs AS ps INNER JOIN tblSubscriptions AS s " & _
        "ON ps.SubID=s.SubID WHERE ps.Status='Pending Approval' ORDER BY ps.RequestedOn, ps.PersonSubID"
End Function

' ---------------------------------------------------------------------------
' Changes by the signed-in person
' ---------------------------------------------------------------------------

' Ask for subscription subId: a tblPersonSubs row, Status 'Pending
' Approval', for the signed-in person (PersonID, and PersonName from
' TempVars PersonName), with the delivery chosen and the justification.
'   freq           Real-time, Daily Digest or Weekly
'   chInApp, chEmail, chTeams   the channels (at least one)
'   justification  why they need it (required; up to 2000 characters)
' Returns "" when the request is made, else what to tell them (a mandatory
' or inactive subscription, one they already have or asked for, a missing
' justification or channel).
' Errors: ERR_NOT_ALLOWED without the Subscription Requester role
'   (RequireRole "REQUEST"); the engine's and KitExecute's, after a rollback.
' Side effects: one tblPersonSubs row and an audit row, in one transaction.
Public Function RequestSubscription(ByVal subId As Long, ByVal freq As String, ByVal chInApp As Boolean, ByVal chEmail As Boolean, _
        ByVal chTeams As Boolean, ByVal justification As String) As String
    Dim pid As Long, st As String, nm As String, ws As DAO.Workspace, inTrans As Boolean
    justification = Trim$(justification)
    If Not DataLinked() Then RequestSubscription = "This copy has no data file, so requests can't be made.": Exit Function
    pid = SignedInPersonID()
    If pid = 0 Then RequestSubscription = "Nobody is signed in.": Exit Function
    nm = Nz(DLookup("SubName", "tblSubscriptions", "SubID=" & subId & " AND IsActive"), "")
    If Len(nm) = 0 Then RequestSubscription = "That subscription is no longer offered.": Exit Function
    If Nz(DLookup("IsMandatory", "tblSubscriptions", "SubID=" & subId), False) Then
        RequestSubscription = nm & " is assigned by command: you already receive it.": Exit Function
    End If
    If freq <> "Real-time" And freq <> "Daily Digest" And freq <> "Weekly" Then
        RequestSubscription = "Pick how often: Real-time, Daily Digest or Weekly.": Exit Function
    End If
    If Not (chInApp Or chEmail Or chTeams) Then RequestSubscription = "Pick at least one way to receive it.": Exit Function
    If Len(justification) = 0 Then RequestSubscription = "Say why you need it: the approver reads this.": Exit Function
    If Len(justification) > 2000 Then RequestSubscription = "Keep the justification to 2000 characters or fewer.": Exit Function
    RequireRole "REQUEST", "request a subscription"
    st = MySubStatus(subId)
    Select Case st
        Case "Active", "Paused": RequestSubscription = "You already have " & nm & ".": Exit Function
        Case "Pending Approval": RequestSubscription = "You already asked for " & nm & "; it waits for an approver.": Exit Function
    End Select
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "INSERT INTO tblPersonSubs (SubID, PersonID, PersonName, Status, Frequency, ChInApp, ChEmail, ChTeams, Justification, " & _
        "RequestedOn, LastUpdated) VALUES (" & subId & ", " & pid & ", " & Q(Left$(SignedInName(), 100)) & ", 'Pending Approval', " & _
        Q(freq) & ", " & IIf(chInApp, "True", "False") & ", " & IIf(chEmail, "True", "False") & ", " & _
        IIf(chTeams, "True", "False") & ", " & Q(justification) & ", Now(), Now())"
    AuditLog "Request Subscription", nm, "Routed to unit Subscription Approver. " & freq & "."
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modSubscriptions.RequestSubscription", subId
End Function

' Withdraw the signed-in person's pending request personSubId: the row is
' removed (it was never approved; the audit trail keeps the record).
' Returns "" when withdrawn, else what to tell them (not theirs, or decided).
' Errors: the engine's and KitExecute's, after a rollback.
Public Function WithdrawSubRequest(ByVal personSubId As Long) As String
    Dim nm As String, ws As DAO.Workspace, inTrans As Boolean
    If Not DataLinked() Then Exit Function
    nm = SubNameOfRow(personSubId)
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If KitExecute("DELETE FROM tblPersonSubs WHERE PersonSubID=" & personSubId & " AND PersonID=" & SignedInPersonID() & _
            " AND Status='Pending Approval'") = 0 Then
        ws.Rollback: inTrans = False
        WithdrawSubRequest = "That request is no longer waiting: an approver may have decided it."
        Exit Function
    End If
    AuditLog "Withdraw Subscription", nm, "Request withdrawn before a decision."
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modSubscriptions.WithdrawSubRequest", personSubId
End Function

' Pause (pause True) or resume the signed-in person's subscription row
' personSubId. Only their own non-mandatory row, Active to pause, Paused to
' resume. Returns "" when changed, else what to tell them.
' Errors: the engine's and KitExecute's, after a rollback.
Public Function SetSubscriptionPaused(ByVal personSubId As Long, ByVal pause As Boolean) As String
    Dim nm As String, ws As DAO.Workspace, inTrans As Boolean, fromSt As String, toSt As String
    If Not DataLinked() Then Exit Function
    If Nz(DLookup("IsMandatory", "tblSubscriptions", "SubID=" & SubIdOfRow(personSubId)), False) Then
        SetSubscriptionPaused = "A subscription assigned by command can't be paused.": Exit Function
    End If
    If pause Then fromSt = "Active": toSt = "Paused" Else fromSt = "Paused": toSt = "Active"
    nm = SubNameOfRow(personSubId)
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If KitExecute("UPDATE tblPersonSubs SET Status=" & Q(toSt) & ", LastUpdated=Now() WHERE PersonSubID=" & personSubId & _
            " AND PersonID=" & SignedInPersonID() & " AND Status=" & Q(fromSt)) = 0 Then
        ws.Rollback: inTrans = False
        SetSubscriptionPaused = "That subscription isn't " & LCase$(fromSt) & " any more. The list now shows where it stands."
        Exit Function
    End If
    AuditLog "Update Subscription", nm, "Subscription " & LCase$(toSt) & "."
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modSubscriptions.SetSubscriptionPaused", personSubId
End Function

' ---------------------------------------------------------------------------
' Decisions (Subscription Approver; called by the Admin page)
' ---------------------------------------------------------------------------

' Approve request personSubId: Status Active, with who decided, when and
' the note. For the Admin page's approvals.
'   personSubId  the tblPersonSubs row (Status 'Pending Approval')
'   decisionNote the approver's note, optional (up to 255 characters)
' Returns "" when approved, else what to tell the approver: the request was
' decided or withdrawn meanwhile, or is their own (someone else decides it).
' Errors: ERR_NOT_ALLOWED without the Subscription Approver role
'   (RequireRole "APPROVER"); the engine's and KitExecute's, after a rollback.
' Side effects: the row updated and an audit row "Approve Subscription",
'   in one transaction.
Public Function ApproveSubscription(ByVal personSubId As Long, ByVal decisionNote As String) As String
    ApproveSubscription = DecideSub(personSubId, True, decisionNote)
End Function

' Deny request personSubId: Status Denied, with who decided, when and why.
'   personSubId  the tblPersonSubs row (Status 'Pending Approval')
'   decisionNote why (required: the requester reads it; up to 255 characters)
' Returns "" when denied, else what to tell the approver (no reason given,
' decided or withdrawn meanwhile, or their own request).
' Errors: ERR_NOT_ALLOWED without the Subscription Approver role
'   (RequireRole "APPROVER"); the engine's and KitExecute's, after a rollback.
' Side effects: the row updated and an audit row "Deny Subscription", in
'   one transaction.
Public Function DenySubscription(ByVal personSubId As Long, ByVal decisionNote As String) As String
    DenySubscription = DecideSub(personSubId, False, decisionNote)
End Function

' ApproveSubscription and DenySubscription.
Private Function DecideSub(ByVal personSubId As Long, ByVal approve As Boolean, ByVal decisionNote As String) As String
    Dim pid As Long, ws As DAO.Workspace, inTrans As Boolean, nm As String, who As String, owner As Variant
    decisionNote = Trim$(Replace(Replace(decisionNote, vbCr, " "), vbLf, " "))
    If Not DataLinked() Then DecideSub = "This copy has no data file.": Exit Function
    If Not approve And Len(decisionNote) = 0 Then DecideSub = "Say why the request is denied: the requester reads it.": Exit Function
    If Len(decisionNote) > 255 Then DecideSub = "Keep the note to 255 characters or fewer.": Exit Function
    RequireRole "APPROVER", IIf(approve, "approve", "deny") & " a subscription request"
    pid = SignedInPersonID()
    owner = DLookup("PersonID", "tblPersonSubs", "PersonSubID=" & personSubId)
    If Not IsNull(owner) Then
        If owner = pid Then DecideSub = "This is your own request: another Subscription Approver decides it.": Exit Function
    End If
    nm = SubNameOfRow(personSubId)
    who = Nz(DLookup("PersonName", "tblPersonSubs", "PersonSubID=" & personSubId), "")
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If KitExecute("UPDATE tblPersonSubs SET Status=" & Q(IIf(approve, "Active", "Denied")) & ", DeciderID=" & IIf(pid = 0, "Null", pid) & _
            ", DeciderName=" & Q(Left$(SignedInName(), 100)) & ", DecidedOn=Now(), LastUpdated=Now(), DecisionNote=" & _
            IIf(Len(decisionNote) = 0, "Null", Q(decisionNote)) & " WHERE PersonSubID=" & personSubId & " AND Status='Pending Approval'") = 0 Then
        ws.Rollback: inTrans = False
        DecideSub = "That request is no longer waiting: it was decided or withdrawn."
        Exit Function
    End If
    AuditLog IIf(approve, "Approve Subscription", "Deny Subscription"), nm & " for " & who, _
        IIf(approve, "Request approved.", "Request denied.") & IIf(Len(decisionNote) > 0, " Note: " & decisionNote, "")
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modSubscriptions.DecideSub", personSubId
End Function

' The subscription name of a tblPersonSubs row, for the audit trail.
Private Function SubNameOfRow(ByVal personSubId As Long) As String
    SubNameOfRow = Nz(DLookup("SubName", "tblSubscriptions", "SubID=" & SubIdOfRow(personSubId)), "subscription request " & personSubId)
End Function

' The SubID of a tblPersonSubs row, 0 when the row is gone.
Private Function SubIdOfRow(ByVal personSubId As Long) As Long
    SubIdOfRow = Nz(DLookup("SubID", "tblPersonSubs", "PersonSubID=" & personSubId), 0)
End Function
