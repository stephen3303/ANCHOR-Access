Option Compare Database
Option Explicit

' ANCHOR (Access edition) - messages: each person's read and acknowledgement
' state, and the SQL every message list shares.
'
' A message is read-only: it is the official record as it arrived. What a
' person does to it lives in tblMessageUserState (no row: unread):
'   ReadState     unread or read (the web app's "new" is an unread message
'                 received since the person's last sign-in, MessageIsNew)
'   FirstReadOn   when they first opened it
'   AckOn         when they acknowledged it (designated traffic only)
' Opening a message marks it read and audits the view (MarkRead);
' Acknowledge records the acknowledgement and audits it. Traffic the
' person's own command sent (Source = tblSettings CommandSource) never needs
' their acknowledgement: the command tracks the addressees' receipts instead
' (tblAckReceipts, the Message page's acknowledgement status).
'
' Pages open a message with OpenMessage, which hands its id to the Message
' page through TempVars OpenMessageID.
'
' Owner: messages. Depends on: modAudit, modErrors, modGlobals, modSignIn;
' the kit's modKitLocks, modPersonnel.

' The Message page's key (frmMain.PageForm).
Public Const PAGE_MESSAGE As String = "message"

' Rank of a precedence, highest first: FLASH 1, IMMEDIATE 2, PRIORITY 3,
' ROUTINE 4, anything else 5. For sorting (PrecedenceRankSql in queries).
Public Function PrecedenceRank(ByVal p As Variant) As Long
    Select Case UCase$(Nz(p, ""))
        Case "FLASH": PrecedenceRank = 1
        Case "IMMEDIATE": PrecedenceRank = 2
        Case "PRIORITY": PrecedenceRank = 3
        Case "ROUTINE": PrecedenceRank = 4
        Case Else: PrecedenceRank = 5
    End Select
End Function

' The same rank as an SQL expression over column col.
Public Function PrecedenceRankSql(Optional ByVal col As String = "m.Precedence") As String
    PrecedenceRankSql = "Switch(" & col & "='FLASH',1," & col & "='IMMEDIATE',2," & col & "='PRIORITY',3," & col & "='ROUTINE',4,True,5)"
End Function

' FROM clause: every message with the signed-in person's state joined as s
' (s.ReadState, s.AckOn Null when they have no row).
Public Function MessagesFromSql() As String
    MessagesFromSql = "tblMessages AS m LEFT JOIN (SELECT MessageID, ReadState, FirstReadOn, AckOn FROM tblMessageUserState " & _
        "WHERE PersonID=" & SignedInPersonID() & ") AS s ON m.MessageID=s.MessageID"
End Function

' SQL expression: True when the signed-in person must still acknowledge the
' message (designated traffic, not their command's own, not acknowledged).
Public Function NeedsAckSql() As String
    NeedsAckSql = "(m.AckRequired AND m.Source<>" & Q(CommandSource()) & " AND s.AckOn Is Null)"
End Function

' SQL expression: "unread" or "read" for the signed-in person.
Public Function ReadStateSql() As String
    ReadStateSql = "IIf(s.ReadState Is Null,'unread',s.ReadState)"
End Function

' The columns the message lists share, for a SELECT over MessagesFromSql:
' id, the read state, precedence (and its rank for sorting), the DTG, the
' reference, subject, originator, type, classification, lifecycle, whether
' an acknowledgement is due from the person (ack_txt), and received.
Public Function MessageListColsSql() As String
    MessageListColsSql = "m.MessageID AS id, " & ReadStateSql() & " AS read_state, " & _
        "IIf(" & ReadStateSql() & "='unread',ChrW(9679),'') AS dot, m.Precedence AS prec, " & _
        PrecedenceRankSql() & " AS prec_rank, m.Dtg AS dtg, m.Reference AS ref, m.Subject AS subject, " & _
        "m.Originator AS originator, m.MsgType AS msg_type, m.Classification AS classif, m.Lifecycle AS lifecycle, " & _
        "IIf(" & NeedsAckSql() & ",'Ack due',IIf(s.AckOn Is Null,'','Acknowledged')) AS ack_txt, " & _
        "m.ReceivedOn AS received, Format(m.ReceivedOn,'dd mmm yy hh:nn') AS received_txt"
End Function

' The person's read state of one message: "unread" or "read".
Public Function ReadStateOf(ByVal mid As Long) As String
    ReadStateOf = Nz(DLookup("ReadState", "tblMessageUserState", "MessageID=" & mid & " AND PersonID=" & SignedInPersonID()), "unread")
End Function

' True when the message arrived since the person last signed in and is unread.
Public Function MessageIsNew(ByVal mid As Long) As Boolean
    Dim seen As Variant
    If ReadStateOf(mid) <> "unread" Then Exit Function
    seen = DLookup("LastSeenOn", "tblPeople", "PersonID=" & SignedInPersonID())
    If IsNull(seen) Then MessageIsNew = True: Exit Function
    MessageIsNew = Nz(DLookup("ReceivedOn", "tblMessages", "MessageID=" & mid), 0) >= seen
End Function

' Mark a message read for the signed-in person and audit the view (the
' first time only). No-op with no data file or nobody signed in.
' Errors: KitExecute's.
Public Sub MarkRead(ByVal mid As Long)
    Dim pid As Long
    If Not DataLinked() Then Exit Sub
    pid = SignedInPersonID()
    If pid = 0 Then Exit Sub
    If DCount("*", "tblMessageUserState", "MessageID=" & mid & " AND PersonID=" & pid) = 0 Then
        KitExecute "INSERT INTO tblMessageUserState (MessageID, PersonID, ReadState, FirstReadOn) VALUES (" & mid & ", " & pid & ", 'read', Now())"
        AuditTry "View", "Message (" & MessageRef(mid) & ")", "Message opened and read.", MessageClassification(mid)
    ElseIf ReadStateOf(mid) <> "read" Then
        KitExecute "UPDATE tblMessageUserState SET ReadState='read' WHERE MessageID=" & mid & " AND PersonID=" & pid
    End If
End Sub

' Mark a message unread again (the list's Mark unread).
Public Sub MarkUnread(ByVal mid As Long)
    Dim pid As Long
    If Not DataLinked() Then Exit Sub
    pid = SignedInPersonID()
    If pid = 0 Then Exit Sub
    KitExecute "UPDATE tblMessageUserState SET ReadState='unread' WHERE MessageID=" & mid & " AND PersonID=" & pid
End Sub

' Mark every message read (Messages' Mark all read). Returns how many changed.
Public Function MarkAllRead() As Long
    Dim pid As Long, n As Long
    If Not DataLinked() Then Exit Function
    pid = SignedInPersonID()
    If pid = 0 Then Exit Function
    n = KitExecute("UPDATE tblMessageUserState SET ReadState='read' WHERE PersonID=" & pid & " AND ReadState<>'read'")
    n = n + KitExecute("INSERT INTO tblMessageUserState (MessageID, PersonID, ReadState, FirstReadOn) SELECT MessageID, " & pid & _
        ", 'read', Now() FROM tblMessages WHERE MessageID NOT IN (SELECT MessageID FROM tblMessageUserState WHERE PersonID=" & pid & ")")
    MarkAllRead = n
End Function

' True when the signed-in person must still acknowledge message mid.
Public Function NeedsOwnAck(ByVal mid As Long) As Boolean
    Dim rs As DAO.Recordset
    If Not DataLinked() Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT " & NeedsAckSql() & " AS due FROM " & MessagesFromSql() & " WHERE m.MessageID=" & mid, dbOpenSnapshot)
    If Not rs.EOF Then NeedsOwnAck = Nz(rs!due, False)
    rs.Close
End Function

' How many messages the signed-in person must still acknowledge.
Public Function PendingAckCount() As Long
    Dim rs As DAO.Recordset
    If Not DataLinked() Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM " & MessagesFromSql() & " WHERE " & NeedsAckSql(), dbOpenSnapshot)
    PendingAckCount = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' How many messages are unread for the signed-in person.
Public Function UnreadCount() As Long
    Dim rs As DAO.Recordset
    If Not DataLinked() Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM " & MessagesFromSql() & " WHERE s.ReadState Is Null OR s.ReadState='unread'", dbOpenSnapshot)
    UnreadCount = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' Record the signed-in person's acknowledgement of message mid, and audit it.
' Errors: 5 when the message needs no acknowledgement from them (not
'   designated traffic, their command's own, or already acknowledged);
'   KitExecute's.
Public Sub Acknowledge(ByVal mid As Long)
    Dim pid As Long, ws As DAO.Workspace, inTrans As Boolean
    pid = SignedInPersonID()
    If pid = 0 Then Err.Raise 5, "modMessages.Acknowledge", "Nobody is signed in."
    If Not NeedsOwnAck(mid) Then Err.Raise 5, "modMessages.Acknowledge", "This message doesn't need your acknowledgement."
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If DCount("*", "tblMessageUserState", "MessageID=" & mid & " AND PersonID=" & pid) = 0 Then
        KitExecute "INSERT INTO tblMessageUserState (MessageID, PersonID, ReadState, FirstReadOn, AckOn) VALUES (" & mid & ", " & pid & _
            ", 'read', Now(), Now())"
    Else
        KitExecute "UPDATE tblMessageUserState SET ReadState='read', AckOn=Now() WHERE MessageID=" & mid & " AND PersonID=" & pid
    End If
    AuditLog "Acknowledge", "Message (" & MessageRef(mid) & ")", "Receipt acknowledged in ANCHOR.", MessageClassification(mid)
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modMessages.Acknowledge", mid
End Sub

' When the signed-in person acknowledged message mid, or Null.
Public Function AckOnOf(ByVal mid As Long) As Variant
    AckOnOf = DLookup("AckOn", "tblMessageUserState", "MessageID=" & mid & " AND PersonID=" & SignedInPersonID())
End Function

' "Due in 5 h", "Due 26 OCT 1830Z", "Overdue by 2 d" for an ack due date.
Public Function AckDueText(ByVal due As Variant) As String
    Dim hrs As Double
    If IsNull(due) Then Exit Function
    hrs = (due - Now()) * 24
    If hrs < 0 Then
        If -hrs < 48 Then AckDueText = "Overdue by " & CLng(-hrs) & " h" Else AckDueText = "Overdue by " & CLng(-hrs / 24) & " d"
    ElseIf hrs < 48 Then
        AckDueText = "Due in " & CLng(hrs) & " h"
    Else
        AckDueText = "Due " & UCase$(Format$(due, "dd mmm hhnn")) & "Z"
    End If
End Function

Public Function MessageRef(ByVal mid As Long) As String
    MessageRef = Nz(DLookup("Reference", "tblMessages", "MessageID=" & mid), "message " & mid)
End Function

Public Function MessageClassification(ByVal mid As Long) As String
    MessageClassification = Nz(DLookup("Classification", "tblMessages", "MessageID=" & mid), "UNCLASSIFIED")
End Function

' The id of the message with reference ref ("COMSUBLANT 061/26"), 0 when
' there is none (a related message ANCHOR never received).
Public Function MessageIdByRef(ByVal ref As String) As Long
    MessageIdByRef = Nz(DLookup("MessageID", "tblMessages", "Reference=" & Q(ref)), 0)
End Function

' Open message mid on the Message page, once the calling event has finished.
Public Function OpenMessage(ByVal mid As Long)
    On Error GoTo Err_Handler
    If mid = 0 Then Exit Function
    TempVars("OpenMessageID") = mid
    NavTo PAGE_MESSAGE
    Exit Function
Err_Handler:
    HandleError "modMessages.OpenMessage", mid
End Function

' The lifecycle banner for a message that is no longer current, or "":
' "Superseded by COMSUBLANT 078/26. Read that message for the current direction."
Public Function LifecycleNotice(ByVal lifecycle As String, ByVal replacedBy As Variant, ByVal relatesTo As Variant) As String
    Select Case UCase$(lifecycle)
        Case "SUPERSEDED"
            LifecycleNotice = "Superseded" & IIf(IsNull(replacedBy), "", " by " & replacedBy) & ". Read that message for the current direction."
        Case "CANCELLED"
            LifecycleNotice = "Cancelled" & IIf(IsNull(replacedBy), "", " by " & replacedBy) & ". Take no action on this message."
        Case "EXPIRED"
            LifecycleNotice = "Expired. Kept for the record; it no longer directs action."
        Case "AMENDED"
            LifecycleNotice = "Amended" & IIf(IsNull(replacedBy), "", " by " & replacedBy) & ". Read it with the amendment."
        Case "CORRECTED"
            LifecycleNotice = "Corrected" & IIf(IsNull(relatesTo), "", ": corrects " & relatesTo) & "."
    End Select
End Function

' ---------------------------------------------------------------------------
' Classification marking (demonstration only)
' ---------------------------------------------------------------------------

' ANCHOR carries UNCLASSIFIED, CUI, CONFIDENTIAL and SECRET, for
' demonstration only (Stephen, 2026-11-05): every data file holds synthetic
' content. Wherever a level above UNCLASSIFIED shows (a list row, the message
' view, a print, a briefing, a draft), its demo annotation shows beside it.
Public Const DEMO_MARKING_NOTE As String = "DEMONSTRATION MARKING ONLY. CONTENT IS UNCLASSIFIED."

' True for a classification above UNCLASSIFIED (anything but UNCLASSIFIED or blank).
Public Function IsMarkedAboveUnclass(ByVal classif As Variant) As Boolean
    Dim c As String
    c = UCase$(Trim$(Nz(classif, "")))
    IsMarkedAboveUnclass = (Len(c) > 0 And c <> "UNCLASSIFIED" And c <> "UNCLASS")
End Function

' The marking line for a banner or a print: "UNCLASSIFIED", or for a higher
' level "SECRET -- DEMONSTRATION MARKING ONLY. CONTENT IS UNCLASSIFIED."
Public Function ClassificationMarking(ByVal classif As Variant) As String
    Dim c As String
    c = UCase$(Trim$(Nz(classif, "")))
    If Len(c) = 0 Then c = "UNCLASSIFIED"
    If IsMarkedAboveUnclass(c) Then
        ClassificationMarking = c & " -- " & DEMO_MARKING_NOTE
    Else
        ClassificationMarking = c
    End If
End Function

' The same as an SQL expression over column col, for list columns.
Public Function ClassificationMarkingSql(Optional ByVal col As String = "m.Classification") As String
    ClassificationMarkingSql = "IIf(Nz(" & col & ",'UNCLASSIFIED') In ('UNCLASSIFIED','UNCLASS',''),Nz(" & col & ",'UNCLASSIFIED')," & _
        col & " & ' (demo marking)')"
End Function
