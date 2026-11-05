Option Compare Database
Option Explicit

' ANCHOR (Access edition) - briefings: a title, a period, the messages a
' Briefing Creator picked from that period's traffic with a note on any of
' them, and notes on the whole briefing (tblBriefings, tblBriefingItems).
' Printing one opens rptBriefing (modBuildBriefings.Build_rptBriefing).
'
' While a briefing is being built its pick list lives in a local table of
' this front end, ws_brief_pick (id = MessageID, ticked, ItemNote): "Add
' matching messages" adds the period's messages that match the filters,
' unticked; the creator ticks the ones to brief. Saving writes the ticked
' ones only, in precedence order, newest first.
'
' What a briefing says about a message is the creator's note, or else the
' first paragraph of the official text (the first section, or the body up to
' its first blank line). It never uses the AI summary, which is not the
' official record.
'
' Classification: a briefing is marked with the highest classification of
' its messages (UNCLASSIFIED, CUI, CONFIDENTIAL, SECRET), with the demo
' annotation beside any level above UNCLASSIFIED (ClassificationMarking).
'
' Owner: the Briefings page. Depends on: modAudit, modErrors, modGlobals,
' modMessages, modSignIn; the kit's modKitLocks.

Public Const BRIEFING_REPORT As String = "rptBriefing"

Private Const PICK_TABLE As String = "ws_brief_pick"
Private Const ALL_KEY As String = "*"           ' a filter combo's "all" row

' ---------------------------------------------------------------------------
' Filters
' ---------------------------------------------------------------------------

' The precedence filter's rows, a two-column value list: key, caption.
Public Function BrfPrecedenceList() As String
    BrfPrecedenceList = ValueItem(ALL_KEY) & ";" & ValueItem("All precedences") & ";FLASH;FLASH;IMMEDIATE;IMMEDIATE;" & _
        "PRIORITY;PRIORITY;ROUTINE;ROUTINE"
End Function

' The message type filter's rows, from tblMessageTypes (active types, in
' their order), a two-column value list: key, caption.
' Errors: the engine's.
Public Function BrfTypeList() As String
    Dim rs As DAO.Recordset, s As String
    s = ValueItem(ALL_KEY) & ";" & ValueItem("All types")
    Set rs = CurrentDb.OpenRecordset("SELECT TypeCode FROM tblMessageTypes WHERE IsActive ORDER BY SortOrder, TypeCode", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & ";" & ValueItem(rs!TypeCode) & ";" & ValueItem(rs!TypeCode)
        rs.MoveNext
    Loop
    rs.Close
    BrfTypeList = s
End Function

' One value list item, quoted so a ";" or "," in it stays in it.
Private Function ValueItem(ByVal s As Variant) As String
    ValueItem = """" & Replace(Nz(s, ""), """", """""") & """"
End Function

' The filters as tblBriefings.Criteria: "prec=FLASH|type=*|community=Safety".
Public Function BrfCriteriaText(ByVal prec As String, ByVal msgType As String, ByVal community As String) As String
    BrfCriteriaText = "prec=" & prec & "|type=" & msgType & "|community=" & Replace(community, "|", " ")
End Function

' One filter from a Criteria text (BrfCriteriaText), dflt when it isn't there.
Public Function BrfCriteriaPart(ByVal criteria As String, ByVal key As String, Optional ByVal dflt As String = "") As String
    Dim part As Variant
    BrfCriteriaPart = dflt
    For Each part In Split(criteria, "|")
        If LCase$(Left$(part, Len(key) + 1)) = LCase$(key) & "=" Then BrfCriteriaPart = Mid$(part, Len(key) + 2)
    Next
End Function

' ---------------------------------------------------------------------------
' The pick list
' ---------------------------------------------------------------------------

' Make the pick list's local table when this front end lacks it.
Private Sub EnsurePickTable()
    If Not TableExists(PICK_TABLE, True) Then
        CurrentDb.Execute "CREATE TABLE " & PICK_TABLE & " (id LONG CONSTRAINT pk_ws_brief_pick PRIMARY KEY, ticked YESNO, " & _
            "ItemNote MEMO)", dbFailOnError
    End If
End Sub

' Empty the pick list (a new briefing).
' Errors: the engine's.
Public Sub BrfClearPick()
    EnsurePickTable
    CurrentDb.Execute "DELETE FROM " & PICK_TABLE, dbFailOnError
End Sub

' Fill the pick list with saved briefing briefId's messages, ticked, with
' their notes.
' Errors: the engine's.
Public Sub BrfLoadPick(ByVal briefId As Long)
    BrfClearPick
    CurrentDb.Execute "INSERT INTO " & PICK_TABLE & " (id, ticked, ItemNote) SELECT i.MessageID, True, i.ItemNote " & _
        "FROM tblBriefingItems AS i WHERE i.BriefingID=" & briefId, dbFailOnError
End Sub

' Add the messages received from fromDay to toDay (whole days, both
' included) that match the filters and aren't listed yet, unticked.
'   prec, msgType   a precedence or type, or "*" for any
'   community       text the message's communities contain, "" for any
' Returns how many were added.
' Errors: the engine's.
Public Function BrfAddMatching(ByVal fromDay As Date, ByVal toDay As Date, ByVal prec As String, ByVal msgType As String, _
        ByVal community As String) As Long
    Dim db As DAO.Database, where As String
    EnsurePickTable
    where = "m.ReceivedOn>=" & SqlDay(fromDay) & " AND m.ReceivedOn<" & SqlDay(toDay + 1)
    If Len(prec) > 0 And prec <> ALL_KEY Then where = where & " AND m.Precedence=" & Q(prec)
    If Len(msgType) > 0 And msgType <> ALL_KEY Then where = where & " AND m.MsgType=" & Q(msgType)
    community = Trim$(community)
    If Len(community) > 0 Then where = where & " AND m.Communities Like " & Q("*" & LikeLiteral(community) & "*")
    Set db = CurrentDb
    db.Execute "INSERT INTO " & PICK_TABLE & " (id, ticked, ItemNote) SELECT m.MessageID, False, Null FROM tblMessages AS m " & _
        "WHERE " & where & " AND m.MessageID NOT IN (SELECT id FROM " & PICK_TABLE & ")", dbFailOnError
    BrfAddMatching = db.RecordsAffected
End Function

' A date as an ACE SQL literal, #mm/dd/yyyy#, whatever the regional settings.
Private Function SqlDay(ByVal d As Date) As String
    SqlDay = Format$(DateValue(d), "\#mm\/dd\/yyyy\#")
End Function

' The pick list's rows for its table: ticked mark, reference, DTG,
' precedence, subject (with the demo annotation beside a classification
' above UNCLASSIFIED) and the item's note; highest precedence and newest
' first.
Public Function BrfPickSql() As String
    BrfPickSql = "SELECT p.id, IIf(p.ticked,ChrW(10003),'') AS tick, m.Reference AS ref, m.Dtg AS dtg, m.Precedence AS prec, " & _
        PrecedenceRankSql() & " AS prec_rank, m.Subject & " & MarkingSuffixSql("m.Classification") & " AS subject_txt, " & _
        "Nz(p.ItemNote,'') AS item_note, m.ReceivedOn AS received FROM " & PICK_TABLE & " AS p INNER JOIN tblMessages AS m " & _
        "ON p.id=m.MessageID ORDER BY " & PrecedenceRankSql() & ", m.ReceivedOn DESC"
End Function

' SQL text added after a subject: "" for UNCLASSIFIED, else
' "  [SECRET -- DEMONSTRATION MARKING ONLY. CONTENT IS UNCLASSIFIED.]".
Private Function MarkingSuffixSql(ByVal col As String) As String
    MarkingSuffixSql = "IIf(Nz(" & col & ",'UNCLASSIFIED') In ('UNCLASSIFIED','UNCLASS',''),'','  [' & " & col & " & " & _
        Q(" -- " & DEMO_MARKING_NOTE & "]") & ")"
End Function

' How many messages are listed, and how many of them are ticked.
' Errors: the engine's.
Public Sub BrfPickCounts(ByRef nListed As Long, ByRef nTicked As Long)
    EnsurePickTable
    nListed = DCount("*", PICK_TABLE)
    nTicked = DCount("*", PICK_TABLE, "ticked")
End Sub

' Tick message msgId on the pick list, or untick it.
' Errors: the engine's.
Public Sub BrfToggle(ByVal msgId As Long)
    CurrentDb.Execute "UPDATE " & PICK_TABLE & " SET ticked=Not ticked WHERE id=" & msgId, dbFailOnError
End Sub

' Message msgId's note on the pick list ("" for none).
Public Function BrfNoteOf(ByVal msgId As Long) As String
    BrfNoteOf = Nz(DLookup("ItemNote", PICK_TABLE, "id=" & msgId), "")
End Function

' Set message msgId's note ("" clears it). A message given a note is ticked:
' a note is only worth writing on a message being briefed.
' Errors: the engine's.
Public Sub BrfSetNote(ByVal msgId As Long, ByVal noteText As String)
    noteText = Trim$(noteText)
    CurrentDb.Execute "UPDATE " & PICK_TABLE & " SET ItemNote=" & IIf(Len(noteText) = 0, "Null", Q(noteText)) & _
        IIf(Len(noteText) > 0, ", ticked=True", "") & " WHERE id=" & msgId, dbFailOnError
End Sub

' ---------------------------------------------------------------------------
' Saved briefings
' ---------------------------------------------------------------------------

' The saved briefings, newest first, for the list.
Public Function BrfSavedSql() As String
    BrfSavedSql = "SELECT b.BriefingID AS id, b.Title AS brief_title, Format(b.CreatedOn,'dd mmm yy') AS created_txt, " & _
        "b.CreatedOn FROM tblBriefings AS b ORDER BY b.CreatedOn DESC, b.BriefingID DESC"
End Function

' Save the briefing: a new one when briefId is 0, else replace briefing
' briefId's details and messages with these. The messages are the ticked
' ones on the pick list. Returns the briefing's id.
'   title     required, cut to 150 characters
'   fromDay, toDay   the period
'   criteria  the filters (BrfCriteriaText)
'   notes     notes on the whole briefing, "" for none
' Errors: ERR_NOT_ALLOWED without the Briefing Creator role; 5 with no title
'   or nothing ticked; the engine's and KitExecute's. Nothing is written
'   unless all of it is.
' Side effects: one audit row, in the same transaction.
Public Function BrfSave(ByVal briefId As Long, ByVal title As String, ByVal fromDay As Date, ByVal toDay As Date, _
        ByVal criteria As String, ByVal notes As String) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, rs As DAO.Recordset, n As Long, isNew As Boolean
    title = Trim$(title)
    If Len(title) = 0 Then Err.Raise 5, "modBriefings.BrfSave", "Give the briefing a title."
    EnsurePickTable
    If DCount("*", PICK_TABLE, "ticked") = 0 Then Err.Raise 5, "modBriefings.BrfSave", "Tick at least one message to brief."
    RequireRole "BRIEF", "save briefings"
    isNew = (briefId = 0)
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If isNew Then
        briefId = InsertRow("INSERT INTO tblBriefings (Title, PersonID, PersonName, CreatedOn, PeriodFrom, PeriodTo, Criteria, Notes) " & _
            "VALUES (" & Q(Left$(title, 150)) & ", " & SignedInPersonID() & ", " & Q(Left$(SignedInName(), 100)) & ", Now(), " & _
            SqlDay(fromDay) & ", " & SqlDay(toDay) & ", " & Q(criteria) & ", " & IIf(Len(Trim$(notes)) = 0, "Null", Q(notes)) & ")")
    Else
        If KitExecute("UPDATE tblBriefings SET Title=" & Q(Left$(title, 150)) & ", PeriodFrom=" & SqlDay(fromDay) & ", PeriodTo=" & _
                SqlDay(toDay) & ", Criteria=" & Q(criteria) & ", Notes=" & IIf(Len(Trim$(notes)) = 0, "Null", Q(notes)) & _
                " WHERE BriefingID=" & briefId) = 0 Then
            Err.Raise 5, "modBriefings.BrfSave", "This briefing was deleted since you opened it. Save it again as a new one."
        End If
        KitExecute "DELETE FROM tblBriefingItems WHERE BriefingID=" & briefId
    End If
    ' The ticked messages in briefing order.
    Set rs = CurrentDb.OpenRecordset("SELECT p.id, p.ItemNote FROM " & PICK_TABLE & " AS p INNER JOIN tblMessages AS m " & _
        "ON p.id=m.MessageID WHERE p.ticked ORDER BY " & PrecedenceRankSql() & ", m.ReceivedOn DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        n = n + 1
        KitExecute "INSERT INTO tblBriefingItems (BriefingID, MessageID, SortOrder, ItemNote) VALUES (" & briefId & ", " & rs!id & _
            ", " & n & ", " & IIf(IsNull(rs!ItemNote.Value), "Null", Q(rs!ItemNote.Value)) & ")"
        rs.MoveNext
    Loop
    rs.Close
    AuditLog IIf(isNew, "Create Briefing", "Update Briefing"), "Briefing (" & Left$(title, 120) & ")", _
        "Briefing " & IIf(isNew, "created", "saved") & " with " & n & " " & Plural(n, "message") & " from " & _
        Format$(fromDay, "dd mmm yyyy") & " to " & Format$(toDay, "dd mmm yyyy") & ".", BrfHighestClass(briefId)
    ws.CommitTrans: inTrans = False
    BrfSave = briefId
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modBriefings.BrfSave", title
End Function

' The highest classification among briefing briefId's messages:
' UNCLASSIFIED, CUI, CONFIDENTIAL or SECRET.
' Errors: the engine's.
Public Function BrfHighestClass(ByVal briefId As Long) As String
    Dim rs As DAO.Recordset, r As Long
    Set rs = CurrentDb.OpenRecordset("SELECT Max(Switch(m.Classification='SECRET',3,m.Classification='CONFIDENTIAL',2," & _
        "m.Classification='CUI',1,True,0)) AS r FROM tblBriefingItems AS i INNER JOIN tblMessages AS m " & _
        "ON i.MessageID=m.MessageID WHERE i.BriefingID=" & briefId, dbOpenSnapshot)
    If Not rs.EOF Then r = Nz(rs!r, 0)
    rs.Close
    BrfHighestClass = Choose(r + 1, "UNCLASSIFIED", "CUI", "CONFIDENTIAL", "SECRET")
End Function

' rptBriefing's classification line, top and bottom of each page: the
' highest classification of the briefing's messages, with the demo
' annotation beside a level above UNCLASSIFIED. Called by the report.
Public Function BrfMarking(ByVal briefId As Variant) As String
    On Error GoTo Fail
    BrfMarking = ClassificationMarking(BrfHighestClass(Nz(briefId, 0)))
    Exit Function
Fail:
    LogError "modBriefings.BrfMarking", Nz(briefId, "")
    BrfMarking = "CLASSIFICATION NOT KNOWN: CHECK EACH MESSAGE"
End Function

' What rptBriefing says about one message: "Note: <the creator's note>", or
' "From the official text: <its first paragraph>". Never the AI summary.
' Called by the report.
Public Function BrfItemText(ByVal msgId As Variant, ByVal itemNote As Variant) As String
    On Error GoTo Fail
    If Len(Trim$(Nz(itemNote, ""))) > 0 Then
        BrfItemText = "Note: " & Trim$(itemNote)
    Else
        BrfItemText = "From the official text: " & FirstParagraph(Nz(msgId, 0))
    End If
    Exit Function
Fail:
    LogError "modBriefings.BrfItemText", Nz(msgId, "")
    BrfItemText = "(The official text could not be read: open the message in ANCHOR.)"
End Function

' The first paragraph of message msgId's official text: its first section's
' body, or else its body text up to the first blank line.
Private Function FirstParagraph(ByVal msgId As Long) As String
    Dim rs As DAO.Recordset, s As String, p As Long
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 Body FROM tblMessageSections WHERE MessageID=" & msgId & _
        " ORDER BY SortOrder, SectionID", dbOpenSnapshot)
    If Not rs.EOF Then s = Trim$(Nz(rs!Body, ""))
    rs.Close
    If Len(s) = 0 Then
        s = Replace(Nz(DLookup("BodyText", "tblMessages", "MessageID=" & msgId), ""), vbCrLf, vbLf)
        s = Trim$(Replace(s, vbCr, vbLf))
        p = InStr(s, vbLf & vbLf)
        If p > 0 Then s = Left$(s, p - 1)
        s = Trim$(Replace(s, vbLf, " "))
    End If
    If Len(s) = 0 Then s = "(no text recorded)"
    FirstParagraph = s
End Function

' Print briefing briefId: rptBriefing in preview, and a Print audit row.
' Errors: Access's when the report can't open.
Public Sub BrfPrint(ByVal briefId As Long)
    Dim title As String
    title = Nz(DLookup("Title", "tblBriefings", "BriefingID=" & briefId), "briefing " & briefId)
    On Error Resume Next                    ' one left open from before
    DoCmd.Close acReport, BRIEFING_REPORT, acSaveNo
    On Error GoTo 0
    DoCmd.OpenReport BRIEFING_REPORT, acViewPreview, , "BriefingID=" & briefId
    AuditTry "Print", "Briefing (" & Left$(title, 120) & ")", "Briefing printed (preview opened).", BrfHighestClass(briefId)
End Sub
