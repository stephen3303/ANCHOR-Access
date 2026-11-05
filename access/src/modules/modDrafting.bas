Option Compare Database
Option Explicit

' ANCHOR (Access edition) - drafting and release: outgoing command traffic,
' from a draft through its routing chain to the release that adds it to
' ANCHOR's Messages. Drafting came over from Message Generator (Stephen,
' 2026-11-05); the screens are the web app's Drafting & Release page
' (/drafting) and the rules its lib/routing.ts.
'
' A draft (tblDrafts, its paragraphs in tblDraftParagraphs) moves through:
'   drafting     not yet submitted; only its drafter edits it
'   in_routing   held by the route's current step; the text is frozen
'   returned     a reviewer sent it back with a comment; the drafter edits
'                and resubmits it, which builds a new route
'   released     released by a release authority: DTG, serial and local
'                sequence assigned in release order, the rendered text
'                frozen in RenderedText, and the message added to
'                tblMessages, so it shows in ANCHOR's Messages
'   withdrawn    withdrawn by its drafter before release, kept for the record
'
' Submitting copies the workflow that matches the draft (modRouting
' WorkflowFor) into tblDraftRoute as route RouteNo: each step with who holds
' it, the first one current. The holder of the current step endorses it on
' to the next step or returns it; the holder of the release step releases it
' when they pass the release gate (modRouting.ReleaseGateProblem).
'
' Each action is checked in the data file at the moment of the write: an
' UPDATE that names the state the page showed (the draft's status, the
' step's status and holder) inside one transaction, so two people acting at
' once can't both succeed; the second is told someone acted first. Every
' action adds a row to tblRoutingEvents, which the app only ever appends
' to, and is audited (modAudit).
'
' The message as it transmits (RenderDraftText) follows NTP 3(J) Annex A
' and the NAV26237 sample, as Message Generator's GENADMIN template does:
' the classification line, precedence, prosign and DTG, FM, TO and INFO, BT,
' the text classification, MSGID/GENADMIN, SUBJ, RMKS with numbered
' paragraphs, BT, the local sequence, NNNN and the classification line.
' Lines wrap at 69 characters (NTP 3 A102b). A draft classified above
' UNCLASSIFIED carries Stephen's test-message line after the text
' classification, because every ANCHOR data file is synthetic (modMessages
' DEMO_MARKING_NOTE). The naval format check (modNavalFormat) runs on a
' trial copy with the DTG and serial the draft would take if released now:
' its blocking issues stop a submit and a release; its warnings never do.
'
' Owner: drafting and release. Depends on: modAudit, modErrors, modGlobals,
' modMessages, modNavalFormat, modRouting, modSignIn; the kit's modCrypto
' (Sha256Hex), modKitLocks (KitExecute), modPersonnel.

' A draft's status (tblDrafts.Status).
Public Const DS_DRAFTING As String = "drafting"
Public Const DS_ROUTING As String = "in_routing"
Public Const DS_RETURNED As String = "returned"
Public Const DS_RELEASED As String = "released"
Public Const DS_WITHDRAWN As String = "withdrawn"

' The Drafting page's chips.
Public Const DRAFT_CHIPS As String = "All|Waiting on me|My drafts|Released"

Private Const LINE_WIDTH As Long = 69                ' NTP 3 A102b
Private Const TEST_MARKING As String = "UNCLASSIFIED - CLASSIFICATION MARKINGS FOR TESTING PURPOSES ONLY"
Private Const NO_DTG As String = "------Z --- --"

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

' ---------------------------------------------------------------------------
' Statuses and lists
' ---------------------------------------------------------------------------

' The word a status shows as: Drafting, In Routing, Returned, Released,
' Withdrawn (pills: modUI.StatusVariant).
Public Function DraftStatusWord(ByVal code As Variant) As String
    Select Case Nz(code, "")
        Case DS_DRAFTING: DraftStatusWord = "Drafting"
        Case DS_ROUTING: DraftStatusWord = "In Routing"
        Case DS_RETURNED: DraftStatusWord = "Returned"
        Case DS_RELEASED: DraftStatusWord = "Released"
        Case DS_WITHDRAWN: DraftStatusWord = "Withdrawn"
    End Select
End Function

' One sentence on what a status means, as the web app's hints.
Public Function DraftStatusHint(ByVal code As Variant) As String
    Select Case Nz(code, "")
        Case DS_DRAFTING: DraftStatusHint = "Not yet submitted. Only its drafter edits it."
        Case DS_ROUTING: DraftStatusHint = "Held by a reviewer. The drafter can't change the text."
        Case DS_RETURNED: DraftStatusHint = "Sent back for correction with a reviewer's comment."
        Case DS_RELEASED: DraftStatusHint = "Released by the release authority. The record is frozen."
        Case DS_WITHDRAWN: DraftStatusHint = "Withdrawn before release and kept for the record."
    End Select
End Function

Private Function StatusWordSql(ByVal col As String) As String
    StatusWordSql = "Switch(" & col & "='" & DS_DRAFTING & "','Drafting'," & col & "='" & DS_ROUTING & "','In Routing'," & _
        col & "='" & DS_RETURNED & "','Returned'," & col & "='" & DS_RELEASED & "','Released',True,'Withdrawn')"
End Function

' FROM clause of the draft list: each draft with its current step (c, Null
' when no step is current).
Private Function DraftFromSql() As String
    DraftFromSql = "tblDrafts AS d LEFT JOIN (SELECT DraftID, PersonID AS cur_pid, PersonName AS cur_name, StepTitle AS cur_title " & _
        "FROM tblDraftRoute WHERE StepStatus='current') AS c ON d.DraftID=c.DraftID"
End Function

' WHERE clause for a chip (DRAFT_CHIPS, 0 All) and the search text.
Private Function DraftWhereSql(ByVal chip As Long, ByVal search As String) As String
    Dim w As String, pid As Long
    pid = SignedInPersonID()
    Select Case chip
        Case 1: w = "d.Status='" & DS_ROUTING & "' AND c.cur_pid=" & pid
        Case 2: w = "d.DrafterID=" & pid
        Case 3: w = "d.Status='" & DS_RELEASED & "'"
    End Select
    search = Trim$(search)
    If Len(search) > 0 Then
        w = w & IIf(Len(w) > 0, " AND ", "") & "(Nz(d.Subject,'') & ' ' & Nz(d.MsgType,'') & ' ' & Nz(d.LocalSequence,'') & ' ' & " & _
            "Nz(d.DrafterName,'') & ' ' & Nz(c.cur_name,'')) Like " & Q("*" & LikeLiteral(search) & "*")
    End If
    If Len(w) > 0 Then DraftWhereSql = " WHERE " & w
End Function

' Rows of the Drafting page's list, the latest change first: id, the status
' word ("Your action" when it waits on the signed-in person), precedence,
' subject, type, classification with its demo marking, and who it is with.
'   chip    DRAFT_CHIPS: 0 All, 1 Waiting on me, 2 My drafts, 3 Released
'   search  words in the subject, type, sequence, drafter or holder
Public Function DraftListSql(ByVal chip As Long, ByVal search As String) As String
    DraftListSql = "SELECT d.DraftID AS id, IIf(d.Status='" & DS_ROUTING & "' AND c.cur_pid=" & SignedInPersonID() & ",'Your action'," & _
        StatusWordSql("d.Status") & ") AS status, d.Precedence AS prec, IIf(Nz(d.Subject,'')='','Untitled draft',d.Subject) AS subject, " & _
        "d.MsgType, " & ClassificationMarkingSql("d.Classification") & " AS classif, " & _
        "IIf(d.Status='" & DS_RELEASED & "',d.LocalSequence & ' ' & d.Dtg,IIf(c.cur_name Is Null,'drafted by ' & d.DrafterName," & _
        "'with ' & c.cur_name)) AS where_txt, d.UpdatedOn FROM " & DraftFromSql() & DraftWhereSql(chip, search) & " ORDER BY d.UpdatedOn DESC"
End Function

' How many drafts a chip lists (no search).
Public Function DraftCount(ByVal chip As Long) As Long
    Dim rs As DAO.Recordset
    If Not DataLinked() Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM " & DraftFromSql() & DraftWhereSql(chip, ""), dbOpenSnapshot)
    DraftCount = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' The draft's current route, step by step: id, n, StepTitle, holder,
' its status word and comment.
Public Function RouteStepsSql(ByVal draftId As Long) As String
    RouteStepsSql = "SELECT r.RouteStepID AS id, r.StepOrder AS n, r.StepTitle & IIf(r.StepAction='release',' (release)','') AS step_txt, " & _
        "r.PersonName AS holder, Switch(r.StepStatus='pending','Not yet',r.StepStatus='current','Awaiting action'," & _
        "r.StepStatus='endorsed','Endorsed',r.StepStatus='returned','Returned',r.StepStatus='released','Released',True,'Withdrawn') AS step_state, " & _
        "IIf(r.ActedOn Is Null,'',Format(r.ActedOn,'dd mmm hh:nn')) & IIf(Nz(r.StepComment,'')='','',IIf(r.ActedOn Is Null,'',': ') & r.StepComment) AS acted_txt " & _
        "FROM tblDraftRoute AS r INNER JOIN tblDrafts AS d ON r.DraftID=d.DraftID AND r.RouteNo=d.RouteNo " & _
        "WHERE r.DraftID=" & draftId & " ORDER BY r.StepOrder"
End Function

' The draft's route log, newest first: id, when, what, the step, who, and
' the comment.
Public Function RouteLogSql(ByVal draftId As Long) As String
    RouteLogSql = "SELECT EventID AS id, Format(EventAt,'dd mmm yy hh:nn') AS at_txt, EventAt, " & _
        "Switch(EventType='created','Created',EventType='submitted','Submitted',EventType='advanced','Endorsed',EventType='returned','Returned'," & _
        "EventType='revised','Revised',EventType='released','Released',True,'Withdrawn') & IIf(Nz(StepTitle,'')='','',': ' & StepTitle) AS event_txt, " & _
        "ActorName, EventComment FROM tblRoutingEvents WHERE DraftID=" & draftId & " ORDER BY EventAt DESC, EventID DESC"
End Function

' ---------------------------------------------------------------------------
' A draft's values
' ---------------------------------------------------------------------------

' A new, empty set of draft values with the command's defaults: keys
' subject, msg_type, msg_class, precedence, classification, originator,
' to_line, info_line, paragraphs (the text, a blank line between them).
Public Function NewDraftValues() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d("subject") = ""
    d("msg_type") = DefaultMsgType()
    d("msg_class") = "A"
    d("precedence") = "ROUTINE"
    d("classification") = "UNCLASSIFIED"
    d("originator") = CommandOriginator()
    d("to_line") = ""
    d("info_line") = ""
    d("paragraphs") = ""
    Set NewDraftValues = d
End Function

' GENADMIN when the type list has it, else the first active type.
Private Function DefaultMsgType() As String
    If Not DataLinked() Then DefaultMsgType = "GENADMIN": Exit Function
    If DCount("*", "tblMessageTypes", "TypeCode='GENADMIN' AND IsActive") > 0 Then
        DefaultMsgType = "GENADMIN"
    Else
        DefaultMsgType = Nz(DLookup("TypeCode", "tblMessageTypes", "IsActive AND SortOrder=" & Nz(DMin("SortOrder", "tblMessageTypes", "IsActive"), 0)), "")
    End If
End Function

' Draft draftId's values as NewDraftValues has them, plus id, status,
' drafter_id, drafter_name, route_no, workflow_id, template_id, dtg, msn,
' local_seq, rendered (the frozen text, "" before release), released_on,
' released_by, designation, message_id, created_on, updated_on. Nothing
' when the draft isn't there.
' Errors: the engine's.
Public Function DraftValues(ByVal draftId As Long) As Object
    Dim rs As DAO.Recordset, d As Object
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblDrafts WHERE DraftID=" & draftId, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    Set d = NewDraftValues()
    d("id") = draftId
    d("subject") = Nz(rs!Subject, "")
    d("msg_type") = Nz(rs!MsgType, "")
    d("msg_class") = Nz(rs!MsgClass, "A")
    d("precedence") = Nz(rs!Precedence, "ROUTINE")
    d("classification") = Nz(rs!Classification, "UNCLASSIFIED")
    d("originator") = Nz(rs!Originator, "")
    d("to_line") = Nz(rs!ToLine, "")
    d("info_line") = Nz(rs!InfoLine, "")
    d("status") = Nz(rs!Status, DS_DRAFTING)
    d("drafter_id") = Nz(rs!DrafterID, 0)
    d("drafter_name") = Nz(rs!DrafterName, "")
    d("route_no") = Nz(rs!RouteNo, 0)
    d("workflow_id") = Nz(rs!WorkflowID, 0)
    d("template_id") = Nz(rs!TemplateID, 0)
    d("dtg") = Nz(rs!Dtg, "")
    d("msn") = Nz(rs!Msn, "")
    d("local_seq") = Nz(rs!LocalSequence, "")
    d("rendered") = Nz(rs!RenderedText, "")
    d("released_on") = rs!ReleasedOn
    d("released_by") = Nz(rs!ReleasedByName, "")
    d("designation") = Nz(rs!ReleaseDesignation, "")
    d("message_id") = Nz(rs!MessageID, 0)
    d("created_on") = rs!CreatedOn
    d("updated_on") = rs!UpdatedOn
    rs.Close
    d("paragraphs") = StoredParagraphs(draftId)
    Set DraftValues = d
End Function

' The draft's paragraphs from tblDraftParagraphs, a blank line between them.
Private Function StoredParagraphs(ByVal draftId As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT ParaText FROM tblDraftParagraphs WHERE DraftID=" & draftId & " ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, vbCrLf & vbCrLf, "") & Nz(rs!ParaText, "")
        rs.MoveNext
    Loop
    rs.Close
    StoredParagraphs = s
End Function

' The paragraphs of a draft's text: a blank line (or one of spaces) starts
' the next one. Each keeps its own lines (its subparagraphs); trailing
' spaces and empty paragraphs are left out.
Private Function SplitParagraphs(ByVal text As String) As Collection
    Dim c As New Collection, lines() As String, i As Long, cur As String
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If Len(Trim$(lines(i))) = 0 Then
            If Len(cur) > 0 Then c.Add cur
            cur = ""
        Else
            cur = cur & IIf(Len(cur) > 0, vbCrLf, "") & RTrim$(lines(i))
        End If
    Next
    If Len(cur) > 0 Then c.Add cur
    Set SplitParagraphs = c
End Function

' Why values d can't be saved, "" when they can: lengths the data file
' holds, a known precedence, classification and class.
Private Function ValuesProblem(ByVal d As Object) As String
    If Len(d("subject")) > 255 Then ValuesProblem = "The subject is " & Len(d("subject")) & " characters; keep it to 255.": Exit Function
    If Len(d("to_line")) > 255 Then ValuesProblem = "The action addressees run to " & Len(d("to_line")) & " characters; keep them to 255.": Exit Function
    If Len(d("info_line")) > 255 Then ValuesProblem = "The info addressees run to " & Len(d("info_line")) & " characters; keep them to 255.": Exit Function
    If Len(d("msg_type")) = 0 Then ValuesProblem = "Pick the message type.": Exit Function
    If InStr(";ROUTINE;PRIORITY;IMMEDIATE;FLASH;", ";" & d("precedence") & ";") = 0 Then ValuesProblem = "Pick the precedence.": Exit Function
    If InStr(";UNCLASSIFIED;CUI;CONFIDENTIAL;SECRET;", ";" & d("classification") & ";") = 0 Then
        ValuesProblem = "Pick the classification.": Exit Function
    End If
    If DCount("*", "tblMessageClasses", "ClassCode=" & Q(d("msg_class"))) = 0 Then ValuesProblem = "Pick the message class (A to E).": Exit Function
End Function

' ---------------------------------------------------------------------------
' Making and editing a draft
' ---------------------------------------------------------------------------

' Add a draft for the signed-in person, its drafter, from values d
' (NewDraftValues' keys). Returns its id.
'   templateId  the template it started from, 0 for none
' Errors: 5 with the reason when d can't be saved or nobody is signed in;
'   the engine's.
Public Function CreateDraft(ByVal d As Object, Optional ByVal templateId As Long = 0) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, id As Long, pid As Long, tplName As String
    pid = SignedInPersonID()
    If pid = 0 Then Err.Raise 5, "modDrafting.CreateDraft", "Nobody is signed in."
    why = ValuesProblem(d)
    If Len(why) > 0 Then Err.Raise 5, "modDrafting.CreateDraft", why
    If templateId <> 0 Then tplName = Nz(DLookup("TemplateName", "tblTemplates", "TemplateID=" & templateId), "")
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    id = InsertRow("INSERT INTO tblDrafts (Subject, MsgType, MsgClass, Precedence, Classification, Originator, ToLine, InfoLine, " & _
        "TemplateID, DrafterID, DrafterName, Status, RouteNo, CreatedOn, UpdatedOn) VALUES (" & Q(d("subject")) & ", " & Q(d("msg_type")) & _
        ", " & Q(d("msg_class")) & ", " & Q(d("precedence")) & ", " & Q(d("classification")) & ", " & Q(d("originator")) & ", " & _
        Q(d("to_line")) & ", " & Q(d("info_line")) & ", " & IIf(templateId = 0, "Null", templateId) & ", " & pid & ", " & _
        Q(Left$(SignedInName(), 100)) & ", '" & DS_DRAFTING & "', 0, Now(), Now())")
    WriteParagraphs id, d("paragraphs")
    AddEvent id, "created", "", IIf(Len(tplName) > 0, "Draft opened from the template " & tplName & ".", "Draft opened.")
    AuditLog "Create draft", DraftResource(id, d("subject")), IIf(Len(tplName) > 0, "From the template " & tplName & ".", "New draft."), _
        d("classification")
    ws.CommitTrans: inTrans = False
    CreateDraft = id
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modDrafting.CreateDraft"
End Function

' A new draft with the command's defaults (New draft). Returns its id.
' Errors: CreateDraft's.
Public Function NewBlankDraft() As Long
    NewBlankDraft = CreateDraft(NewDraftValues())
End Function

' Replace the draft's paragraphs with those of text.
Private Sub WriteParagraphs(ByVal draftId As Long, ByVal text As String)
    Dim p As Variant, n As Long, rs As DAO.Recordset
    KitExecute "DELETE FROM tblDraftParagraphs WHERE DraftID=" & draftId
    Set rs = CurrentDb.OpenRecordset("tblDraftParagraphs", dbOpenDynaset, dbAppendOnly)
    For Each p In SplitParagraphs(text)
        n = n + 1
        rs.AddNew
        rs!DraftID = draftId
        rs!SortOrder = n
        rs!ParaText = p
        rs.Update
    Next
    rs.Close
End Sub

' Why the signed-in person can't edit draft draftId now, "" when they can.
' Errors: the engine's.
Public Function DraftEditProblem(ByVal draftId As Long) As String
    Dim st As String, drafter As Long
    st = Nz(DLookup("Status", "tblDrafts", "DraftID=" & draftId), "")
    drafter = Nz(DLookup("DrafterID", "tblDrafts", "DraftID=" & draftId), 0)
    Select Case st
        Case ""
            DraftEditProblem = "This draft is no longer in the data file."
        Case DS_DRAFTING, DS_RETURNED
            If drafter <> SignedInPersonID() Then
                DraftEditProblem = "Only its drafter, " & Nz(DLookup("DrafterName", "tblDrafts", "DraftID=" & draftId), "") & ", can edit it."
            End If
        Case DS_ROUTING
            DraftEditProblem = "The text is frozen while it is in routing. A reviewer returns it for correction."
        Case DS_RELEASED
            DraftEditProblem = "Released: the record is frozen. Changes go out as a corrected copy or an update."
        Case Else
            DraftEditProblem = "Withdrawn before release and kept for the record."
    End Select
End Function

' Save the signed-in person's edits to their draft. Returns "" when saved,
' else why not (to show). A save after a return is logged as Revised once.
' Errors: the engine's.
Public Function SaveDraft(ByVal draftId As Long, ByVal d As Object) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, n As Long, st As String
    why = ValuesProblem(d)
    If Len(why) > 0 Then SaveDraft = why: Exit Function
    why = DraftEditProblem(draftId)
    If Len(why) > 0 Then SaveDraft = why: Exit Function
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    n = KitExecute("UPDATE tblDrafts SET Subject=" & Q(d("subject")) & ", MsgType=" & Q(d("msg_type")) & ", MsgClass=" & Q(d("msg_class")) & _
        ", Precedence=" & Q(d("precedence")) & ", Classification=" & Q(d("classification")) & ", Originator=" & Q(Left$(d("originator"), 60)) & _
        ", ToLine=" & Q(d("to_line")) & _
        ", InfoLine=" & Q(d("info_line")) & ", UpdatedOn=Now() WHERE DraftID=" & draftId & " AND DrafterID=" & SignedInPersonID() & _
        " AND Status IN ('" & DS_DRAFTING & "','" & DS_RETURNED & "')")
    If n <> 1 Then
        ws.Rollback: inTrans = False
        SaveDraft = DraftEditProblem(draftId)
        If Len(SaveDraft) = 0 Then SaveDraft = "The draft changed while you edited it. Pick it again to see it now."
        Exit Function
    End If
    WriteParagraphs draftId, d("paragraphs")
    st = Nz(DLookup("Status", "tblDrafts", "DraftID=" & draftId), "")
    If st = DS_RETURNED And LastEventType(draftId) = "returned" Then AddEvent draftId, "revised", "", "Corrected after the return."
    AuditLog "Change draft", DraftResource(draftId, d("subject")), "Saved.", d("classification")
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modDrafting.SaveDraft", draftId
End Function

' The type of the draft's latest route log entry, "" for none.
Private Function LastEventType(ByVal draftId As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 EventType FROM tblRoutingEvents WHERE DraftID=" & draftId & _
        " ORDER BY EventAt DESC, EventID DESC", dbOpenSnapshot)
    If Not rs.EOF Then LastEventType = Nz(rs!EventType, "")
    rs.Close
End Function

' Append a route log entry for the signed-in person (never changed after).
Private Sub AddEvent(ByVal draftId As Long, ByVal kind As String, ByVal stepTitle As String, ByVal comment As String)
    Dim pid As Long
    pid = SignedInPersonID()
    KitExecute "INSERT INTO tblRoutingEvents (DraftID, EventType, EventAt, ActorID, ActorName, StepTitle, EventComment) VALUES (" & _
        draftId & ", " & Q(kind) & ", Now(), " & IIf(pid = 0, "Null", pid) & ", " & Q(Left$(SignedInName(), 100)) & ", " & _
        IIf(Len(stepTitle) = 0, "Null", Q(Left$(stepTitle, 100))) & ", " & Q(comment) & ")"
End Sub

' What the audit trail calls a draft: "Draft 12 (Holiday leave period watch manning)".
Private Function DraftResource(ByVal draftId As Long, ByVal subject As String) As String
    DraftResource = "Draft " & draftId & IIf(Len(subject) > 0, " (" & ClipText(subject, 80) & ")", "")
End Function

' ---------------------------------------------------------------------------
' The message as it transmits
' ---------------------------------------------------------------------------

' The message text of values d (NewDraftValues' keys), as it transmits.
'   dtg       "041830Z NOV 26", or "" for the placeholder a draft shows
'   serial    the MSGID serial "0042", or "" for "----"
'   localSeq  the line after the closing BT "#0042", or "" for "#----"
' The command's text case (CommandTextCase) applies to everything typed.
' Errors: the engine's (the command settings).
Public Function RenderDraftText(ByVal d As Object, ByVal dtg As String, ByVal serial As String, ByVal localSeq As String) As String
    Dim out As String, up As Boolean, cls As String, prec As String, mon As String, topLine As String
    Dim paras As Collection, i As Long, n As Long, lines() As String, j As Long, s As String, v As Variant, first As Boolean
    up = (CommandTextCase() = "upper")
    cls = UCase$(Nz(d("classification"), ""))
    If Len(cls) = 0 Then cls = "UNCLASSIFIED"
    prec = UCase$(Nz(d("precedence"), ""))
    If Len(prec) = 0 Then prec = "ROUTINE"
    If Len(dtg) = 0 Then dtg = NO_DTG
    mon = Mid$(dtg, 9, 3)
    If Len(serial) = 0 Then serial = NullSerial()
    If Len(localSeq) = 0 Then localSeq = LocalSequenceText(serial)
    topLine = "CLASSIFICATION: " & TopClassText(cls) & "/"

    AddLine out, topLine
    AddLine out, prec
    AddLine out, Left$(prec, 1) & " " & dtg
    AddLine out, "FM " & CaseOf(d("originator"), up)
    first = True
    For Each v In AddressList(d("to_line"))
        AddLine out, IIf(first, "TO ", "") & CaseOf(v, up)
        first = False
    Next
    If first Then AddLine out, "TO"
    first = True
    For Each v In AddressList(d("info_line"))
        AddLine out, IIf(first, "INFO ", "") & CaseOf(v, up)
        first = False
    Next
    AddLine out, "BT"
    AddLine out, TextClassLine(cls)
    If IsMarkedAboveUnclass(cls) Then AddLine out, TEST_MARKING
    AddLine out, "MSGID/GENADMIN/" & MsgIdOriginator() & "/" & serial & "/" & mon & "//"
    s = Trim$(Nz(d("subject"), ""))
    If Len(s) = 0 Then s = "UNTITLED"
    AddLine out, "SUBJ/" & CaseOf(s, up) & "//"

    ' RMKS: one paragraph unnumbered (ACP 121(G) 324j), more numbered 1., 2.
    ' with a blank line between them (NAV26237); the last ends with //.
    Set paras = SplitParagraphs(CaseOf(Nz(d("paragraphs"), ""), up))
    n = paras.Count
    If n = 0 Then AddLine out, "RMKS///"
    For i = 1 To n
        If i > 1 Then AddLine out, ""
        lines = Split(paras(i), vbCrLf)
        For j = 0 To UBound(lines)
            s = lines(j)
            If j = 0 Then
                s = IIf(n > 1, i & ". ", "") & LTrim$(s)
                If i = 1 Then s = "RMKS/" & s
            End If
            If i = n And j = UBound(lines) Then s = s & "//"
            AddLine out, WrapLine(s, LINE_WIDTH)
        Next
    Next
    AddLine out, "BT"
    AddLine out, localSeq
    AddLine out, "NNNN"
    RenderDraftText = out & topLine
End Function

' The text before release: placeholders for the DTG and serial, or the
' frozen text of a released draft.
Public Function DraftPreviewText(ByVal d As Object) As String
    If d.Exists("rendered") Then
        If Len(d("rendered")) > 0 Then DraftPreviewText = d("rendered"): Exit Function
    End If
    DraftPreviewText = RenderDraftText(d, "", "", "")
End Function

' The text with the DTG and serial the draft would take if released now,
' for the format check.
Public Function DraftTrialText(ByVal d As Object) As String
    Dim serial As String
    serial = SerialText(NextSequenceNumber())
    DraftTrialText = RenderDraftText(d, DtgOf(UtcNow()), serial, LocalSequenceText(serial))
End Function

' The naval format check of values d (modNavalFormat), on DraftTrialText.
Public Function DraftFormatIssues(ByVal d As Object) As String
    DraftFormatIssues = NavalFormatIssues(DraftTrialText(d))
End Function

Private Sub AddLine(ByRef out As String, ByVal s As String)
    out = out & s & vbCrLf
End Sub

Private Function CaseOf(ByVal v As Variant, ByVal up As Boolean) As String
    If up Then CaseOf = UCase$(Nz(v, "")) Else CaseOf = Nz(v, "")
End Function

' The addresses of a TO or INFO line, split on commas, semicolons or line
' breaks, trimmed, blanks left out.
Private Function AddressList(ByVal s As Variant) As Collection
    Dim c As New Collection, v As Variant
    For Each v In Split(Replace(Replace(Replace(Replace(Nz(s, ""), vbCrLf, ";"), vbLf, ";"), vbCr, ";"), ",", ";"), ";")
        If Len(Trim$(v)) > 0 Then c.Add Trim$(v)
    Next
    Set AddressList = c
End Function

' The classification line's word: CUI is unclassified, marked //CUI.
Private Function TopClassText(ByVal cls As String) As String
    If cls = "CUI" Then TopClassText = "UNCLASSIFIED//CUI" Else TopClassText = cls
End Function

' The text classification after BT (NTP 3 702a spells a classified one
' with spaces).
Private Function TextClassLine(ByVal cls As String) As String
    Select Case cls
        Case "CUI": TextClassLine = "UNCLAS//CUI"
        Case "CONFIDENTIAL": TextClassLine = "C O N F I D E N T I A L"
        Case "SECRET": TextClassLine = "S E C R E T"
        Case Else: TextClassLine = "UNCLAS"
    End Select
End Function

' One line of text broken at spaces into lines of at most width characters
' (a word longer than that is cut). The first line keeps its indent; the
' others start at the margin, as NAV26237's continuation lines do.
Private Function WrapLine(ByVal s As String, ByVal width As Long) As String
    Dim out As String, p As Long
    Do While Len(s) > width
        p = InStrRev(s, " ", width + 1)
        If p <= 1 Or Len(Trim$(Left$(s, p))) = 0 Then p = width + 1
        out = out & RTrim$(Left$(s, p - 1)) & vbCrLf
        s = LTrim$(Mid$(s, p))
    Loop
    WrapLine = out & s
End Function

' Now in UTC, from Windows.
Public Function UtcNow() As Date
    Dim st As SYSTEMTIME
    GetSystemTime st
    UtcNow = DateSerial(st.wYear, st.wMonth, st.wDay) + TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function

' "041830Z NOV 26" for a UTC time, whatever Windows' language. A time of
' 0000 is written 0001 (NTP 3 504).
Public Function DtgOf(ByVal utc As Date) As String
    Dim hhnn As String
    hhnn = Format$(Hour(utc), "00") & Format$(Minute(utc), "00")
    If hhnn = "0000" Then hhnn = "0001"
    DtgOf = Format$(Day(utc), "00") & hhnn & "Z " & Mid$("JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC", (Month(utc) - 1) * 3 + 1, 3) & _
        " " & Format$(Year(utc) Mod 100, "00")
End Function

' ---------------------------------------------------------------------------
' Routing
' ---------------------------------------------------------------------------

' The draft's current step (tblDraftRoute RouteStepID), 0 when none.
Public Function DraftCurrentStepID(ByVal draftId As Long) As Long
    DraftCurrentStepID = Nz(DLookup("RouteStepID", "tblDraftRoute", "DraftID=" & draftId & " AND StepStatus='current'"), 0)
End Function

' Why the signed-in person can't submit draft draftId (or resubmit it after
' a return), "" when they can: it is theirs and editable, has a subject, an
' action addressee and text, a workflow matches it whose every step someone
' can act on, and the naval format check finds nothing blocking.
'   d   its values (DraftValues)
' Errors: the engine's.
Public Function SubmitProblem(ByVal draftId As Long, ByVal d As Object) As String
    Dim wf As Long, why As String
    why = DraftEditProblem(draftId)
    If Len(why) > 0 Then SubmitProblem = why: Exit Function
    If Len(Trim$(d("subject"))) = 0 Then SubmitProblem = "Give the draft a subject.": Exit Function
    If AddressList(d("to_line")).Count = 0 Then SubmitProblem = "Name at least one action addressee (TO).": Exit Function
    If SplitParagraphs(d("paragraphs")).Count = 0 Then SubmitProblem = "Write the text of the message.": Exit Function
    wf = WorkflowFor(d("msg_type"), d("originator"))
    If wf = 0 Then
        SubmitProblem = "No routing workflow matches " & d("msg_type") & " from " & d("originator") & ". An administrator must define one first."
        Exit Function
    End If
    why = WorkflowProblem(wf)
    If Len(why) > 0 Then SubmitProblem = WorkflowName(wf) & ": " & why: Exit Function
    why = StrandedText(wf)
    If Len(why) > 0 Then SubmitProblem = why & " It can't be submitted until then.": Exit Function
    why = FormatIssuesOf(DraftFormatIssues(d), False)
    If Len(why) > 0 Then SubmitProblem = "The naval format check found problems to fix first:" & vbCrLf & why
End Function

' Submit the signed-in person's draft into the workflow that matches it:
' its steps become route RouteNo + 1, the first one current. Returns "" when
' submitted, else why not.
' Errors: the engine's.
Public Function SubmitDraft(ByVal draftId As Long) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, d As Object, wf As Long, routeNo As Long, n As Long
    Dim rs As DAO.Recordset, k As Long, resub As Boolean
    Set d = DraftValues(draftId)
    If d Is Nothing Then SubmitDraft = "This draft is no longer in the data file.": Exit Function
    why = SubmitProblem(draftId, d)
    If Len(why) > 0 Then SubmitDraft = why: Exit Function
    wf = WorkflowFor(d("msg_type"), d("originator"))
    resub = (d("status") = DS_RETURNED)
    ' The new route's number is read before the transaction: a DLookup inside
    ' it can't see the transaction's own update, so it would read the old
    ' number and file the steps under the previous route (an empty chain on
    ' the Route tab). The update names the number read, so a second submit
    ' at the same moment changes nothing and is refused below.
    routeNo = Nz(DLookup("RouteNo", "tblDrafts", "DraftID=" & draftId), 0) + 1
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    n = KitExecute("UPDATE tblDrafts SET Status='" & DS_ROUTING & "', WorkflowID=" & wf & ", RouteNo=" & routeNo & ", UpdatedOn=Now() " & _
        "WHERE DraftID=" & draftId & " AND DrafterID=" & SignedInPersonID() & " AND Status='" & d("status") & "' AND Nz(RouteNo,0)=" & (routeNo - 1))
    If n <> 1 Then
        ws.Rollback: inTrans = False
        SubmitDraft = "The draft changed before it could be submitted. Pick it again to see it now."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT StepOrder, StepTitle, StepAction, PersonID FROM tblWorkflowSteps WHERE WorkflowID=" & wf & _
        " ORDER BY StepOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        k = k + 1
        KitExecute "INSERT INTO tblDraftRoute (DraftID, RouteNo, StepOrder, StepTitle, StepAction, PersonID, PersonName, StepStatus) VALUES (" & _
            draftId & ", " & routeNo & ", " & k & ", " & Q(rs!StepTitle) & ", " & Q(rs!StepAction) & ", " & rs!PersonID & ", " & _
            Q(Left$(RosterName(rs!PersonID), 100)) & ", " & IIf(k = 1, "'current'", "'pending'") & ")"
        rs.MoveNext
    Loop
    rs.Close
    AddEvent draftId, "submitted", "", IIf(resub, "Resubmitted", "Submitted") & " on " & WorkflowName(wf) & "."
    AuditLog IIf(resub, "Resubmit draft", "Submit draft"), DraftResource(draftId, d("subject")), "Routed on " & WorkflowName(wf) & ": " & _
        WorkflowChainText(wf) & ".", d("classification")
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modDrafting.SubmitDraft", draftId
End Function

' Lock draft draftId's current step for the signed-in person inside the
' open transaction: the step must still be current and theirs, and the
' draft in routing. Returns the step's id, or 0 (with why set) when it isn't.
Private Function TakeCurrentStep(ByVal draftId As Long, ByRef why As String, ByRef stepTitle As String, ByRef stepAction As String) As Long
    Dim rs As DAO.Recordset, n As Long
    n = KitExecute("UPDATE tblDrafts SET UpdatedOn=Now() WHERE DraftID=" & draftId & " AND Status='" & DS_ROUTING & "'")
    If n <> 1 Then why = "This draft is no longer in routing: someone acted on it first. Pick it again to see it now.": Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT RouteStepID, StepTitle, StepAction, PersonID, PersonName FROM tblDraftRoute WHERE DraftID=" & _
        draftId & " AND StepStatus='current'", dbOpenSnapshot)
    If rs.EOF Then
        why = "This draft has no step waiting. Pick it again to see it now."
    ElseIf Nz(rs!PersonID, 0) <> SignedInPersonID() Or SignedInPersonID() = 0 Then
        why = "Only " & Nz(rs!PersonName, "the person holding it") & " can act on " & Nz(rs!StepTitle, "this step") & "."
    Else
        TakeCurrentStep = rs!RouteStepID
        stepTitle = Nz(rs!StepTitle, "")
        stepAction = Nz(rs!StepAction, "review")
    End If
    rs.Close
End Function

' The holder of the current review step endorses it: it moves on to the
' next step. Returns "" or why not.
'   comment   optional, kept on the step and in the route log
' Errors: the engine's.
Public Function EndorseDraft(ByVal draftId As Long, ByVal comment As String) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, stepId As Long, title As String, action As String, nextId As Long, ord As Long
    comment = Trim$(comment)
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    stepId = TakeCurrentStep(draftId, why, title, action)
    If stepId <> 0 And action = "release" Then why = "This is the release step: release the message, or return it for correction."
    If Len(why) > 0 Then
        ws.Rollback: inTrans = False
        EndorseDraft = why
        Exit Function
    End If
    ord = DLookup("StepOrder", "tblDraftRoute", "RouteStepID=" & stepId)
    nextId = Nz(DLookup("RouteStepID", "tblDraftRoute", "DraftID=" & draftId & " AND RouteNo=" & DLookup("RouteNo", "tblDraftRoute", _
        "RouteStepID=" & stepId) & " AND StepOrder=" & (ord + 1)), 0)
    KitExecute "UPDATE tblDraftRoute SET StepStatus='endorsed', ActedOn=Now(), StepComment=" & IIf(Len(comment) = 0, "Null", Q(comment)) & _
        " WHERE RouteStepID=" & stepId & " AND StepStatus='current'"
    If nextId <> 0 Then KitExecute "UPDATE tblDraftRoute SET StepStatus='current' WHERE RouteStepID=" & nextId & " AND StepStatus='pending'"
    AddEvent draftId, "advanced", title, IIf(Len(comment) > 0, comment, "Endorsed.")
    AuditLog "Endorse draft", DraftResource(draftId, Nz(DLookup("Subject", "tblDrafts", "DraftID=" & draftId), "")), _
        title & " endorsed" & IIf(nextId <> 0, "; it waits on " & Nz(DLookup("PersonName", "tblDraftRoute", "RouteStepID=" & nextId), "the next step"), "") & ".", _
        Nz(DLookup("Classification", "tblDrafts", "DraftID=" & draftId), "UNCLASSIFIED")
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modDrafting.EndorseDraft", draftId
End Function

' The holder of the current step returns the draft to its drafter for
' correction. Returns "" or why not.
'   comment   required: what to correct
' Errors: the engine's.
Public Function ReturnDraft(ByVal draftId As Long, ByVal comment As String) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, stepId As Long, title As String, action As String
    comment = Trim$(comment)
    If Len(comment) = 0 Then ReturnDraft = "Type a comment saying what to correct.": Exit Function
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    stepId = TakeCurrentStep(draftId, why, title, action)
    If Len(why) > 0 Then
        ws.Rollback: inTrans = False
        ReturnDraft = why
        Exit Function
    End If
    KitExecute "UPDATE tblDraftRoute SET StepStatus='returned', ActedOn=Now(), StepComment=" & Q(comment) & _
        " WHERE RouteStepID=" & stepId & " AND StepStatus='current'"
    KitExecute "UPDATE tblDrafts SET Status='" & DS_RETURNED & "', UpdatedOn=Now() WHERE DraftID=" & draftId
    AddEvent draftId, "returned", title, comment
    AuditLog "Return draft", DraftResource(draftId, Nz(DLookup("Subject", "tblDrafts", "DraftID=" & draftId), "")), _
        "Returned for correction at " & title & ": " & comment, Nz(DLookup("Classification", "tblDrafts", "DraftID=" & draftId), "UNCLASSIFIED")
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modDrafting.ReturnDraft", draftId
End Function

' Why the signed-in person can't release draft draftId now, "" when they
' can: they hold its current step, the release step, and pass the release
' gate (modRouting.ReleaseGateProblem).
' Errors: the engine's.
Public Function ReleaseProblem(ByVal draftId As Long) As String
    Dim rs As DAO.Recordset, pid As Long
    pid = SignedInPersonID()
    Set rs = CurrentDb.OpenRecordset("SELECT r.StepAction, r.PersonID, r.PersonName, d.MsgType FROM tblDraftRoute AS r INNER JOIN tblDrafts AS d " & _
        "ON r.DraftID=d.DraftID WHERE r.DraftID=" & draftId & " AND r.StepStatus='current' AND d.Status='" & DS_ROUTING & "'", dbOpenSnapshot)
    If rs.EOF Then
        ReleaseProblem = "This draft has no step waiting."
    ElseIf rs!StepAction <> "release" Then
        ReleaseProblem = "It reaches the release authority after its review steps."
    ElseIf Nz(rs!PersonID, 0) <> pid Or pid = 0 Then
        ReleaseProblem = "Only " & Nz(rs!PersonName, "the person holding the release step") & " can release it."
    Else
        ReleaseProblem = ReleaseGateProblem(pid, Nz(rs!MsgType, ""))
    End If
    rs.Close
End Function

' Release the draft: the holder of its release step, a release authority
' for its type, assigns the DTG (now, UTC) and the next serial and local
' sequence, freezes the rendered text, and adds the message to tblMessages
' (Source CommandSource, so ANCHOR's Messages shows it as the command's
' own). All in one transaction with the route log entry and the audit row.
' Returns "" when released, else why not.
'   comment   optional, kept on the step and in the route log
' Errors: the engine's.
Public Function ReleaseDraft(ByVal draftId As Long, ByVal comment As String) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, stepId As Long, title As String, action As String
    Dim d As Object, n As Long, utc As Date, dtg As String, serial As String, seq As String, raw As String, msgId As Long
    Dim authId As Long, desig As String, ref As String, pid As Long
    comment = Trim$(comment)
    pid = SignedInPersonID()
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    stepId = TakeCurrentStep(draftId, why, title, action)
    If Len(why) = 0 Then why = ReleaseProblem(draftId)
    If Len(why) = 0 Then
        Set d = DraftValues(draftId)
        authId = ReleaseAuthorityFor(pid, d("msg_type"))
        desig = ReleaseDesignationOf(authId)
        ' The sequence is taken here, in release order: the update names the
        ' number read, so two releases at once can't take the same one.
        n = NextSequenceNumber()
        If KitExecute("UPDATE tblCommandSettings SET NextSequence=" & (n + 1) & " WHERE SettingsID=1 AND NextSequence=" & n) <> 1 Then
            If DCount("*", "tblCommandSettings") = 0 Then
                why = "The command transmission settings are missing. A System Admin saves them under Admin: Routing."
            Else
                why = "Another release took sequence number " & n & " at the same moment. Release again."
            End If
        End If
    End If
    If Len(why) = 0 Then
        utc = UtcNow()
        dtg = DtgOf(utc)
        serial = SerialText(n)
        seq = LocalSequenceText(serial)
        raw = RenderDraftText(d, dtg, serial, seq)
        why = FormatIssuesOf(NavalFormatIssues(raw), False)
        If Len(why) > 0 Then why = "The naval format check found problems; return the draft for correction:" & vbCrLf & why
    End If
    If Len(why) > 0 Then
        ws.Rollback: inTrans = False
        ReleaseDraft = why
        Exit Function
    End If
    ref = MsgIdOriginator() & " " & serial & "/" & Right$(dtg, 2)
    msgId = AddReleasedMessage(draftId, d, raw, dtg, ref, utc)
    KitExecute "UPDATE tblDraftRoute SET StepStatus='released', ActedOn=Now(), StepComment=" & IIf(Len(comment) = 0, "Null", Q(comment)) & _
        " WHERE RouteStepID=" & stepId & " AND StepStatus='current'"
    KitExecute "UPDATE tblDrafts SET Status='" & DS_RELEASED & "', Dtg=" & Q(dtg) & ", Msn=" & Q(serial) & ", LocalSequence=" & Q(seq) & _
        ", ReleasedOn=Now(), ReleasedByID=" & pid & ", ReleasedByName=" & Q(Left$(SignedInName(), 100)) & ", ReleaseDesignation=" & _
        Q(Left$(desig, 100)) & ", MessageID=" & msgId & ", UpdatedOn=Now() WHERE DraftID=" & draftId
    SetRenderedText draftId, raw
    AddEvent draftId, "released", title, "Released as " & desig & " - " & seq & " - " & dtg & ". The transmitted text is frozen." & _
        IIf(Len(comment) > 0, " " & comment, "")
    AuditLog "Release", "Message (" & ref & ")", "Released as " & desig & ", " & seq & ", " & dtg & ", from " & DraftResource(draftId, d("subject")) & _
        ". Added to Messages; SHA-256 " & Left$(Sha256Hex(raw), 16) & ChrW(8230), d("classification")
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modDrafting.ReleaseDraft", draftId
End Function

' Freeze the released text on the draft (a memo, so not through SQL text).
Private Sub SetRenderedText(ByVal draftId As Long, ByVal raw As String)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT RenderedText FROM tblDrafts WHERE DraftID=" & draftId, dbOpenDynaset)
    rs.Edit
    rs!RenderedText = raw
    rs.Update
    rs.Close
End Sub

' Add the released message to tblMessages, with its sections and
' addressees, as ANCHOR stores received traffic: RawText the text as it
' transmits (verbatim, hashed), BodyText and the sections as typed, a
' local C2OIX id. Returns the new MessageID.
Private Function AddReleasedMessage(ByVal draftId As Long, ByVal d As Object, ByVal raw As String, ByVal dtg As String, _
        ByVal ref As String, ByVal utc As Date) As Long
    Dim rs As DAO.Recordset, paras As Collection, i As Long, body As String, head As String, txt As String, v As Variant, msgRow As Long
    Set paras = SplitParagraphs(d("paragraphs"))
    For i = 1 To paras.Count
        body = body & IIf(i > 1, vbCrLf & vbCrLf, "") & IIf(paras.Count > 1, i & ". ", "") & paras(i)
    Next
    Set rs = CurrentDb.OpenRecordset("tblMessages", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!MsgKey = LCase$("local-" & Replace(Replace(ref, " ", "-"), "/", "-")) & "-d" & draftId
    rs!Dtg = dtg
    rs!ReceivedOn = utc
    rs!Subject = Left$(IIf(Len(Trim$(d("subject"))) = 0, "Untitled", Trim$(d("subject"))), 255)
    rs!Reference = Left$(ref, 60)
    rs!MsgType = d("msg_type")
    rs!Originator = Left$(d("originator"), 60)
    rs!Source = Left$(CommandSource(), 40)
    rs!ToLine = Left$(JoinAddresses(d("to_line")), 255)
    rs!InfoLine = Left$(JoinAddresses(d("info_line")), 255)
    rs!Precedence = d("precedence")
    rs!Classification = d("classification")
    rs!Lifecycle = "ACTIVE"
    rs!AckRequired = False
    rs!C2oixId = Left$("LOCAL-" & Replace(Replace(ref, " ", "-"), "/", "-"), 30)
    rs!Sha256 = Sha256Hex(raw)
    rs!SizeKb = (Len(raw) + 1023) \ 1024
    rs!MsgClass = d("msg_class")
    rs!RawText = raw
    rs!BodyText = body
    rs.Update
    rs.Bookmark = rs.LastModified
    msgRow = rs!MessageID
    rs.Close
    Set rs = CurrentDb.OpenRecordset("tblMessageSections", dbOpenDynaset, dbAppendOnly)
    For i = 1 To paras.Count
        SectionParts paras(i), i, head, txt
        rs.AddNew
        rs!MessageID = msgRow
        rs!SortOrder = i
        rs!Heading = Left$(head, 100)
        rs!Body = txt
        rs.Update
    Next
    rs.Close
    For Each v In AddressList(d("to_line"))
        AddAddressee msgRow, CStr(v), "TO"
    Next
    For Each v In AddressList(d("info_line"))
        AddAddressee msgRow, CStr(v), "INFO"
    Next
    AddReleasedMessage = msgRow
End Function

Private Sub AddAddressee(ByVal msgRow As Long, ByVal nm As String, ByVal role As String)
    KitExecute "INSERT INTO tblAddressees (MessageID, AddrName, AddrRole, AddrKind) VALUES (" & msgRow & ", " & Q(Left$(nm, 100)) & ", " & _
        Q(role) & ", " & IIf(DCount("*", "tblAddressGroups", "GroupName=" & Q(nm)) > 0, "'group'", "'command'") & ")"
End Sub

Private Function JoinAddresses(ByVal s As Variant) As String
    Dim v As Variant, out As String
    For Each v In AddressList(s)
        out = out & IIf(Len(out) > 0, ", ", "") & v
    Next
    JoinAddresses = out
End Function

' A paragraph as a message section: "PURPOSE. To establish ..." gives the
' heading Purpose and the text after it; any other paragraph is
' "Paragraph n" with all its text.
Private Sub SectionParts(ByVal p As String, ByVal n As Long, ByRef head As String, ByRef txt As String)
    Dim firstLine As String, k As Long, w As String
    firstLine = Split(p, vbCrLf)(0)
    k = InStr(firstLine, ". ")
    If k = 0 And Right$(firstLine, 1) = "." Then k = Len(firstLine)
    If k > 1 And k <= 40 Then
        w = Left$(firstLine, k - 1)
        If StrComp(w, UCase$(w), vbBinaryCompare) = 0 And w Like "*[A-Z]*" And Not w Like "*[!A-Z ]*" Then
            head = StrConv(w, vbProperCase)
            txt = Trim$(Mid$(p, k + 1))
            Exit Sub
        End If
    End If
    head = "Paragraph " & n
    txt = p
End Sub

' The drafter withdraws their draft before release (from routing too):
' kept for the record. Returns "" or why not.
' Errors: the engine's.
Public Function WithdrawDraft(ByVal draftId As Long) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, n As Long
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    n = KitExecute("UPDATE tblDrafts SET Status='" & DS_WITHDRAWN & "', UpdatedOn=Now() WHERE DraftID=" & draftId & " AND DrafterID=" & _
        SignedInPersonID() & " AND Status IN ('" & DS_DRAFTING & "','" & DS_ROUTING & "','" & DS_RETURNED & "')")
    If n <> 1 Then
        ws.Rollback: inTrans = False
        WithdrawDraft = "Only its drafter can withdraw a draft, and only before it is released."
        Exit Function
    End If
    KitExecute "UPDATE tblDraftRoute SET StepStatus='withdrawn' WHERE DraftID=" & draftId & " AND StepStatus='current'"
    AddEvent draftId, "withdrawn", "", "Withdrawn before release and kept for the record."
    AuditLog "Withdraw draft", DraftResource(draftId, Nz(DLookup("Subject", "tblDrafts", "DraftID=" & draftId), "")), _
        "Withdrawn before release.", Nz(DLookup("Classification", "tblDrafts", "DraftID=" & draftId), "UNCLASSIFIED")
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modDrafting.WithdrawDraft", draftId
End Function
