Option Compare Database
Option Explicit

' ANCHOR (Access edition) - what the Home, Messages and Message pages and
' the printed message (rptMessage) read: the SQL behind their tables, the
' acknowledgement status of traffic this command sent, the related
' messages, and the text blocks of the official rendition.
'
' Read state, acknowledgements and the shared list columns stay in
' modMessages; this module only reads, it never writes.
'
' Acknowledgement status (the web app's resolveAckStatus): a sent message's
' addressees are expanded to commands, a group addressee (AddrKind 'group')
' to the commands tblGroupMembers lists for it (an unknown group gives
' none), any other addressee to itself. Each command is acknowledged when
' tblAckReceipts holds a receipt from it for the message.
'
' Owner: the message pages' build task. Depends on: modGlobals (Q,
' LikeLiteral), modMessages, modSignIn (CommandSource, SignedInName,
' SettingText); the kit's modPersonnel.

' The list filters the Messages page passes to MessagesListSql.
Public Const MSG_SHOW_ALL As String = "all"
Public Const MSG_SHOW_UNREAD As String = "unread"
Public Const MSG_SHOW_ACK As String = "ack"

' The first entry of each Messages filter combo, which means "no filter".
Public Const MSG_ALL_PREC As String = "All precedences"
Public Const MSG_ALL_TYPES As String = "All types"
Public Const MSG_ALL_CLASSES As String = "All classifications"

' Lifecycle words of a message that no longer directs action.
Private Const RETIRED_SQL As String = "('SUPERSEDED','CANCELLED','EXPIRED')"

' ---------------------------------------------------------------------------
' Messages page
' ---------------------------------------------------------------------------

' The Messages page's rows, newest first: MessageListColsSql plus ack_due
' ("Ack due" or ""), narrowed by the filters.
'   show     MSG_SHOW_ALL, MSG_SHOW_UNREAD or MSG_SHOW_ACK
'   prec     a precedence, or MSG_ALL_PREC / "" for any
'   msgType  a message type code, or MSG_ALL_TYPES / "" for any
'   classif  a classification, or MSG_ALL_CLASSES / "" for any
'   search   text the subject, reference or originator must contain ("" none)
' Errors: none of its own.
Public Function MessagesListSql(ByVal show As String, ByVal prec As String, ByVal msgType As String, _
        ByVal classif As String, ByVal search As String) As String
    Dim w As String, pat As String
    Select Case show
        Case MSG_SHOW_UNREAD: w = w & " AND (s.ReadState Is Null OR s.ReadState='unread')"
        Case MSG_SHOW_ACK: w = w & " AND " & NeedsAckSql()
    End Select
    If Len(prec) > 0 And prec <> MSG_ALL_PREC Then w = w & " AND m.Precedence=" & Q(prec)
    If Len(msgType) > 0 And msgType <> MSG_ALL_TYPES Then w = w & " AND m.MsgType=" & Q(msgType)
    If Len(classif) > 0 And classif <> MSG_ALL_CLASSES Then w = w & " AND m.Classification=" & Q(classif)
    search = Trim$(search)
    If Len(search) > 0 Then
        pat = Q("*" & LikeLiteral(search) & "*")
        w = w & " AND (m.Subject Like " & pat & " OR m.Reference Like " & pat & " OR m.Originator Like " & pat & ")"
    End If
    MessagesListSql = "SELECT " & MessageListColsSql() & ", IIf(" & NeedsAckSql() & ",'Ack due','') AS ack_due FROM " & _
        MessagesFromSql() & IIf(Len(w) > 0, " WHERE " & Mid$(w, 6), "") & " ORDER BY m.ReceivedOn DESC, m.MessageID DESC"
End Function

' The message type combo's value list: MSG_ALL_TYPES, then the active
' types of tblMessageTypes in their order, each quoted. Only
' MSG_ALL_TYPES with no data file.
Public Function MessageTypeValueList() As String
    Dim rs As DAO.Recordset, s As String
    s = VLItem(MSG_ALL_TYPES)
    If DataLinked() Then
        Set rs = CurrentDb.OpenRecordset("SELECT TypeCode FROM tblMessageTypes WHERE IsActive ORDER BY SortOrder, TypeCode", dbOpenSnapshot)
        Do While Not rs.EOF
            s = s & ";" & VLItem(rs!TypeCode)
            rs.MoveNext
        Loop
        rs.Close
    End If
    MessageTypeValueList = s
End Function

' One value list item, in double quotes, so a ";" or "," in it is kept.
Private Function VLItem(ByVal v As Variant) As String
    VLItem = """" & Replace(Nz(v, ""), """", """""") & """"
End Function

' ---------------------------------------------------------------------------
' Home page
' ---------------------------------------------------------------------------

' The messages the signed-in person must still acknowledge, the soonest due
' first (no due date last): id, subject, ref, prec, prec_rank, due_txt
' (AckDueText) and due.
Public Function HomeAckSql() As String
    HomeAckSql = "SELECT m.MessageID AS id, m.Subject AS subject, m.Reference AS ref, m.Precedence AS prec, " & _
        PrecedenceRankSql() & " AS prec_rank, AckDueText(m.AckDue) AS due_txt, m.AckDue AS due FROM " & MessagesFromSql() & _
        " WHERE " & NeedsAckSql() & " ORDER BY IIf(m.AckDue Is Null,1,0), m.AckDue, m.ReceivedOn DESC"
End Function

' The messages whose lifecycle is not ACTIVE, newest first: id, lifecycle,
' and item_txt (the reference, and on a second line what replaced it or what
' it updates).
Public Function HomeStatusSql() As String
    HomeStatusSql = "SELECT m.MessageID AS id, m.Lifecycle AS lifecycle, m.Reference & " & _
        "IIf(m.ReplacedBy Is Null,IIf(m.RelatesTo Is Null,'',Chr(13) & Chr(10) & 'Updates ' & m.RelatesTo)," & _
        "Chr(13) & Chr(10) & 'Replaced by ' & m.ReplacedBy) AS item_txt, m.ReceivedOn AS received " & _
        "FROM tblMessages AS m WHERE m.Lifecycle<>'ACTIVE' ORDER BY m.ReceivedOn DESC, m.MessageID DESC"
End Function

' The n newest messages: id, subject, ref, dtg, prec, prec_rank, received.
Public Function HomeLatestSql(ByVal n As Long) As String
    HomeLatestSql = "SELECT TOP " & n & " m.MessageID AS id, m.Subject AS subject, m.Reference AS ref, m.Dtg AS dtg, " & _
        "m.Precedence AS prec, " & PrecedenceRankSql() & " AS prec_rank, m.ReceivedOn AS received " & _
        "FROM tblMessages AS m ORDER BY m.ReceivedOn DESC, m.MessageID DESC"
End Function

' How many messages carry an action item that still stands (the message is
' not superseded, cancelled or expired). 0 with no data file.
Public Function OpenActionCount() As Long
    If Not DataLinked() Then Exit Function
    OpenActionCount = DCount("*", "tblMessages", "ActionText Is Not Null AND Lifecycle Not In " & RETIRED_SQL)
End Function

' How many messages are no longer ACTIVE (amended, corrected, superseded,
' cancelled or expired). 0 with no data file.
Public Function StatusChangeCount() As Long
    If Not DataLinked() Then Exit Function
    StatusChangeCount = DCount("*", "tblMessages", "Lifecycle<>'ACTIVE'")
End Function

' The designated traffic the signed-in person must acknowledge, and how many
' of those they have: total and done (ByRef). Their command's own traffic
' is not counted (NeedsAckSql). Both 0 with no data file.
Public Sub OwnAckProgress(ByRef total As Long, ByRef done As Long)
    Dim rs As DAO.Recordset
    total = 0: done = 0
    If Not DataLinked() Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) AS n, Sum(IIf(s.AckOn Is Null,0,1)) AS n_done FROM " & MessagesFromSql() & _
        " WHERE m.AckRequired AND m.Source<>" & Q(CommandSource()), dbOpenSnapshot)
    If Not rs.EOF Then
        total = Nz(rs!n, 0)
        done = Nz(rs!n_done, 0)
    End If
    rs.Close
End Sub

' ---------------------------------------------------------------------------
' Message page
' ---------------------------------------------------------------------------

' True when message msgId is traffic this command sent (Source = CommandSource).
Public Function SentByOwnCommand(ByVal msgId As Long) As Boolean
    If Not DataLinked() Then Exit Function
    SentByOwnCommand = (Nz(DLookup("Source", "tblMessages", "MessageID=" & msgId), "") = CommandSource())
End Function

' The commands message msgId was addressed to, groups expanded, each with
' whether it has acknowledged: id and cmd_name (the command's name),
' ack_state ("Acknowledged" or "Pending"), ack_at (its first receipt,
' "dd mmm hh:nn", or "") and first_at. Pending commands first.
Public Function AckStatusSql(ByVal msgId As Long) As String
    AckStatusSql = AckStatusCore(msgId) & " ORDER BY IIf(r.first_at Is Null,0,1), u.cmd"
End Function

' AckStatusSql without its ORDER BY, to count over.
Private Function AckStatusCore(ByVal msgId As Long) As String
    Dim cmds As String
    cmds = "SELECT gm.CommandName AS cmd FROM (tblAddressees AS a INNER JOIN tblAddressGroups AS ag ON a.AddrName=ag.GroupName) " & _
        "INNER JOIN tblGroupMembers AS gm ON ag.GroupID=gm.GroupID WHERE a.MessageID=" & msgId & " AND a.AddrKind='group' " & _
        "UNION SELECT a.AddrName AS cmd FROM tblAddressees AS a WHERE a.MessageID=" & msgId & " AND a.AddrKind<>'group'"
    AckStatusCore = "SELECT u.cmd AS id, u.cmd AS cmd_name, IIf(r.first_at Is Null,'Pending','Acknowledged') AS ack_state, " & _
        "IIf(r.first_at Is Null,'',Format(r.first_at,'dd mmm hh:nn')) AS ack_at, r.first_at FROM (" & cmds & ") AS u " & _
        "LEFT JOIN (SELECT CommandName, Min(ReceivedOn) AS first_at FROM tblAckReceipts WHERE MessageID=" & msgId & _
        " GROUP BY CommandName) AS r ON u.cmd=r.CommandName"
End Function

' How many addressee commands message msgId has (groups expanded) and how
' many of them have acknowledged it: total and acked (ByRef).
Public Sub AckStatusCounts(ByVal msgId As Long, ByRef total As Long, ByRef acked As Long)
    Dim rs As DAO.Recordset
    total = 0: acked = 0
    If Not DataLinked() Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) AS n, Sum(IIf(q.ack_state='Acknowledged',1,0)) AS n_acked FROM (" & _
        AckStatusCore(msgId) & ") AS q", dbOpenSnapshot)
    If Not rs.EOF Then
        total = Nz(rs!n, 0)
        acked = Nz(rs!n_acked, 0)
    End If
    rs.Close
End Sub

' The messages ANCHOR holds that message msgId names: its Related list
' ("REF A; REF B"), what it relates to and what replaced it. Columns id,
' lifecycle, item_txt (reference, then the subject on a second line), received.
' Returns "" when it names none.
Public Function RelatedMessagesSql(ByVal msgId As Long) As String
    Dim rs As DAO.Recordset, refs As String, part As Variant, own As String
    Set rs = CurrentDb.OpenRecordset("SELECT Reference, Related, RelatesTo, ReplacedBy FROM tblMessages WHERE MessageID=" & msgId, dbOpenSnapshot)
    If Not rs.EOF Then
        own = Nz(rs!Reference, "")
        For Each part In Split(Nz(rs!Related, "") & ";" & Nz(rs!RelatesTo, "") & ";" & Nz(rs!ReplacedBy, ""), ";")
            part = Trim$(part)
            If Len(part) > 0 And part <> own Then refs = refs & "," & Q(part)
        Next
    End If
    rs.Close
    If Len(refs) = 0 Then Exit Function
    RelatedMessagesSql = "SELECT m.MessageID AS id, m.Lifecycle AS lifecycle, m.Reference & Chr(13) & Chr(10) & m.Subject AS item_txt, " & _
        "m.ReceivedOn AS received FROM tblMessages AS m WHERE m.MessageID<>" & msgId & " AND m.Reference In (" & Mid$(refs, 2) & _
        ") ORDER BY m.ReceivedOn DESC"
End Function

' The communications class line for a message's MsgClass: "A - Official
' Department of Defense", the code alone when tblMessageClasses lacks it,
' "" when there is none.
Public Function MsgClassText(ByVal cls As Variant) As String
    Dim nm As Variant
    If Len(Nz(cls, "")) = 0 Then Exit Function
    nm = DLookup("ClassName", "tblMessageClasses", "ClassCode=" & Q(cls))
    MsgClassText = cls & IIf(IsNull(nm), "", " - " & nm)
End Function

' Text with bare line feeds (or carriage returns) turned into CR LF, so a
' text box shows its lines. Text already in CR LF is returned as it is.
Public Function ShowLines(ByVal s As Variant) As String
    Dim t As String
    t = Nz(s, "")
    If InStr(t, vbCrLf) > 0 Then ShowLines = t: Exit Function
    t = Replace(t, vbCr, vbLf)
    ShowLines = Replace(t, vbLf, vbCrLf)
End Function

' ---------------------------------------------------------------------------
' The printed rendition (rptMessage)
' ---------------------------------------------------------------------------

' The heading block of message msgId's official rendition, as the web app's
' print view sets it out: a status line when it no longer directs action,
' precedence and DTG, FM, TO, INFO, BT, the classification marking, MSGID
' and SUBJ.
Public Function MessagePrintHead(ByVal msgId As Variant) As String
    Dim rs As DAO.Recordset, s As String
    If IsNull(msgId) Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblMessages WHERE MessageID=" & CLng(msgId), dbOpenSnapshot)
    If Not rs.EOF Then
        Select Case UCase$(Nz(rs!Lifecycle, ""))
            Case "SUPERSEDED", "CANCELLED", "EXPIRED"
                s = "*** " & UCase$(rs!Lifecycle) & IIf(IsNull(rs!ReplacedBy), "", " BY " & rs!ReplacedBy) & _
                    ". NOT THE CURRENT DIRECTION. ***" & vbCrLf & vbCrLf
        End Select
        s = s & Nz(rs!Precedence, "") & " " & Nz(rs!Dtg, "") & vbCrLf & _
            "FM " & Nz(rs!Originator, "") & vbCrLf & _
            "TO " & Nz(rs!ToLine, "") & vbCrLf & _
            "INFO " & Nz(rs!InfoLine, "") & vbCrLf & _
            "BT" & vbCrLf & _
            ClassificationMarking(rs!Classification) & vbCrLf & _
            "MSGID/" & Nz(rs!MsgType, "") & "/" & Nz(rs!Reference, "") & "//" & vbCrLf & _
            "SUBJ/" & UCase$(Nz(rs!Subject, "")) & "//"
    End If
    rs.Close
    MessagePrintHead = s
End Function

' The integrity block under message msgId's rendition: status, source and
' C2OIX id, SHA-256, class, size, retention, received, who it was printed
' for and when, and the notes that a printed copy is uncontrolled and the
' data is fictional.
Public Function MessagePrintIntegrity(ByVal msgId As Variant) As String
    Dim rs As DAO.Recordset, s As String
    If IsNull(msgId) Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblMessages WHERE MessageID=" & CLng(msgId), dbOpenSnapshot)
    If Not rs.EOF Then
        s = "ANCHOR INTEGRITY VALIDATION" & vbCrLf & _
            "Status: " & Nz(rs!Lifecycle, "") & IIf(IsNull(rs!ReplacedBy), "", " (see " & rs!ReplacedBy & ")") & vbCrLf & _
            "Source: C2OIX " & Nz(rs!C2oixId, "(no id recorded)") & vbCrLf & _
            "SHA-256: " & Nz(rs!Sha256, "(none recorded)") & vbCrLf & _
            "Verification: the SHA-256 of the raw text recorded at C2OIX ingestion." & vbCrLf
        If Len(Nz(rs!MsgClass, "")) > 0 Then s = s & "Class: " & MsgClassText(rs!MsgClass) & vbCrLf
        s = s & "Size: " & Nz(rs!SizeKb, 0) & " KB. Retention: " & Nz(rs!Retention, "not set") & "." & vbCrLf & _
            "Received: " & Format$(rs!ReceivedOn, "dd mmm yyyy hh:nn") & vbCrLf & _
            "Rendered for " & SignedInName() & ", " & Format$(Now(), "dd mmm yyyy hh:nn") & vbCrLf & _
            "Printed copies are uncontrolled. Verify the current version in ANCHOR before acting." & vbCrLf & _
            "DEMONSTRATION: fictional traffic, not an official record."
    End If
    rs.Close
    MessagePrintIntegrity = s
End Function
