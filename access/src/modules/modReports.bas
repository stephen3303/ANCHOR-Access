Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Reports page's figures: acknowledgement
' accountability and traffic measures over a period (the last 7 or 30 days,
' or all traffic).
'
' Accountability covers every message in the period that must be
' acknowledged (AckRequired), except one superseded or cancelled, since it no
' longer directs action. Who must acknowledge depends on who sent it:
'   received traffic   every active ANCHOR person (tblPeople IsActive), each
'                      acknowledging for themselves (tblMessageUserState AckOn)
'   our own traffic    (Source = the command's CommandSource) the addressee
'                      commands, a group (tblAddressees AddrKind 'group')
'                      expanded to its member commands (tblGroupMembers);
'                      a command has acknowledged once tblAckReceipts holds a
'                      receipt from it. A command reached as TO and as INFO
'                      counts once, as TO.
' A message is Acknowledged when everyone it counts has acknowledged, Overdue
' when its AckDue has passed before that, else In Progress.
'
' RptRun works the table out into local tables of this front end, which the
' page's tables read and Export writes out:
'   ws_rptack      one row per message (id = MessageID)
'   ws_rptackdet   the people or commands of the message picked on the page
'
' Owner: the Reports page. Depends on: modAudit, modGlobals, modMessages,
' modSignIn; the kit's modCsv, modKitFiles, modKitUI, modPersonnel.

' The periods the page offers (Segmented keys) and the measures.
Public Const RPT_PERIOD_7 As String = "d7"
Public Const RPT_PERIOD_30 As String = "d30"
Public Const RPT_PERIOD_ALL As String = "all"

Private Const ACK_TABLE As String = "ws_rptack"
Private Const DET_TABLE As String = "ws_rptackdet"
Private Const BASIS_PEOPLE As String = "Our people"
Private Const BASIS_COMMANDS As String = "Addressees"

' ---------------------------------------------------------------------------
' The period
' ---------------------------------------------------------------------------

' SQL condition on m.ReceivedOn for a period key (RPT_PERIOD_*): the last 7
' or 30 days up to now, or "1=1" for all traffic.
Public Function RptPeriodSql(ByVal periodKey As String) As String
    Select Case periodKey
        Case RPT_PERIOD_7: RptPeriodSql = "m.ReceivedOn>=" & SqlDateTime(Now() - 7)
        Case RPT_PERIOD_30: RptPeriodSql = "m.ReceivedOn>=" & SqlDateTime(Now() - 30)
        Case Else: RptPeriodSql = "1=1"
    End Select
End Function

' The period in words, for the page and the export: "Received since 06 Oct 2026".
Public Function RptPeriodText(ByVal periodKey As String) As String
    Select Case periodKey
        Case RPT_PERIOD_7: RptPeriodText = "Received in the last 7 days, since " & Format$(Now() - 7, "dd mmm yyyy hh:nn")
        Case RPT_PERIOD_30: RptPeriodText = "Received in the last 30 days, since " & Format$(Now() - 30, "dd mmm yyyy hh:nn")
        Case Else: RptPeriodText = "All traffic received"
    End Select
End Function

' A date and time as an ACE SQL literal, #mm/dd/yyyy hh:nn:ss# whatever the
' computer's regional settings.
Private Function SqlDateTime(ByVal d As Date) As String
    SqlDateTime = Format$(d, "\#mm\/dd\/yyyy hh\:nn\:ss\#")
End Function

' How many messages arrived in the period.
' Errors: the engine's.
Public Function RptMessageCount(ByVal periodKey As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM tblMessages AS m WHERE " & RptPeriodSql(periodKey), dbOpenSnapshot)
    RptMessageCount = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' ---------------------------------------------------------------------------
' Accountability
' ---------------------------------------------------------------------------

' Make the local tables when this front end lacks them.
' Errors: the engine's. Side effects: creates tables in the front end.
Private Sub EnsureTables()
    If Not TableExists(ACK_TABLE, True) Then
        CurrentDb.Execute "CREATE TABLE " & ACK_TABLE & " (id LONG CONSTRAINT pk_ws_rptack PRIMARY KEY, ref TEXT(60), " & _
            "subject TEXT(255), prec TEXT(12), prec_rank LONG, classif TEXT(20), classif_txt TEXT(100), due DATETIME, " & _
            "due_txt TEXT(60), basis TEXT(20), acked LONG, total LONG, ack_txt TEXT(20), status TEXT(20), received DATETIME)", _
            dbFailOnError
    End If
    If Not TableExists(DET_TABLE, True) Then
        CurrentDb.Execute "CREATE TABLE " & DET_TABLE & " (id LONG CONSTRAINT pk_ws_rptackdet PRIMARY KEY, MessageID LONG, " & _
            "who TEXT(150), via TEXT(255), addr_role TEXT(10), status TEXT(20), ack_on DATETIME, when_txt TEXT(60))", dbFailOnError
    End If
End Sub

' Work out the accountability table for the period into ws_rptack (see the
' top), and empty ws_rptackdet. Returns how many messages it lists.
' Errors: the engine's.
' Side effects: rewrites the two local tables.
Public Function RptRun(ByVal periodKey As String) As Long
    Dim db As DAO.Database, rs As DAO.Recordset, own As String, nPeople As Long, acked As Object
    Dim msgId As Long, a As Long, t As Long, basis As String, n As Long, cmds As Object, k As Variant, v As Variant
    Set db = CurrentDb
    EnsureTables
    db.Execute "DELETE FROM " & ACK_TABLE, dbFailOnError
    db.Execute "DELETE FROM " & DET_TABLE, dbFailOnError
    own = CommandSource()
    nPeople = DCount("*", "tblPeople", "IsActive")
    ' How many active people acknowledged each message, in one pass.
    Set acked = CreateObject("Scripting.Dictionary")
    Set rs = db.OpenRecordset("SELECT s.MessageID, Count(*) AS n FROM tblMessageUserState AS s INNER JOIN tblPeople AS p " & _
        "ON s.PersonID=p.PersonID WHERE p.IsActive AND s.AckOn Is Not Null GROUP BY s.MessageID", dbOpenSnapshot)
    Do While Not rs.EOF
        acked(CLng(rs!MessageID)) = CLng(rs!n)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT m.MessageID, m.Reference, m.Subject, m.Precedence, m.Classification, m.AckDue, m.Source, " & _
        "m.ReceivedOn FROM tblMessages AS m WHERE m.AckRequired AND m.Lifecycle Not In ('SUPERSEDED','CANCELLED') AND " & _
        RptPeriodSql(periodKey), dbOpenSnapshot)
    Do While Not rs.EOF
        msgId = rs!MessageID
        If Nz(rs!Source.Value, "") = own Then
            basis = BASIS_COMMANDS
            Set cmds = AddresseeCommands(msgId)
            t = cmds.Count: a = 0
            For Each k In cmds.Keys
                v = cmds(k)
                If Not IsNull(v(2)) Then a = a + 1
            Next
        Else
            basis = BASIS_PEOPLE
            t = nPeople
            a = 0
            If acked.Exists(msgId) Then a = acked(msgId)
        End If
        db.Execute "INSERT INTO " & ACK_TABLE & " (id, ref, subject, prec, prec_rank, classif, classif_txt, due, due_txt, basis, " & _
            "acked, total, ack_txt, status, received) VALUES (" & msgId & ", " & Q(rs!Reference.Value) & ", " & Q(Left$(Nz(rs!Subject.Value, ""), 255)) & _
            ", " & Q(rs!Precedence.Value) & ", " & PrecedenceRank(rs!Precedence.Value) & ", " & Q(rs!Classification.Value) & ", " & _
            Q(Left$(ClassificationMarking(rs!Classification.Value), 100)) & ", " & SqlDateOrNull(rs!AckDue.Value) & ", " & _
            Q(DueText(rs!AckDue.Value)) & ", " & Q(basis) & ", " & a & ", " & t & ", " & Q(a & " of " & t) & ", " & _
            Q(AckStatus(a, t, rs!AckDue.Value)) & ", " & SqlDateOrNull(rs!ReceivedOn.Value) & ")", dbFailOnError
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    RptRun = n
End Function

' Acknowledged, Overdue or In Progress (StatusVariant words), for a acked of
' t and the due date.
Private Function AckStatus(ByVal a As Long, ByVal t As Long, ByVal due As Variant) As String
    If t > 0 And a >= t Then
        AckStatus = "Acknowledged"
    ElseIf Not IsNull(due) Then
        If due < Now() Then AckStatus = "Overdue" Else AckStatus = "In Progress"
    Else
        AckStatus = "In Progress"
    End If
End Function

' "26 Oct 26 18:30", or "No due time".
Private Function DueText(ByVal due As Variant) As String
    If IsNull(due) Then DueText = "No due time" Else DueText = Format$(due, "dd mmm yy hh:nn")
End Function

Private Function SqlDateOrNull(ByVal d As Variant) As String
    If IsNull(d) Then SqlDateOrNull = "Null" Else SqlDateOrNull = SqlDateTime(CDate(d))
End Function

' The commands a message of ours was sent to, groups expanded to their
' members: a Dictionary of command name -> Array(role TO or INFO, the groups
' it was reached through "a, b" or "", the first receipt's time or Null, how
' many receipts).
' Errors: the engine's.
Private Function AddresseeCommands(ByVal msgId As Long) As Object
    Dim d As Object, rs As DAO.Recordset, mem As DAO.Recordset, nm As String, rl As String, grp As String, cur As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set rs = CurrentDb.OpenRecordset("SELECT AddrName, AddrRole, AddrKind FROM tblAddressees WHERE MessageID=" & msgId & _
        " ORDER BY AddresseeID", dbOpenSnapshot)
    Do While Not rs.EOF
        rl = UCase$(Nz(rs!AddrRole, "TO"))
        If LCase$(Nz(rs!AddrKind, "")) = "group" Then
            grp = rs!AddrName
            Set mem = CurrentDb.OpenRecordset("SELECT gm.CommandName FROM tblGroupMembers AS gm INNER JOIN tblAddressGroups AS g " & _
                "ON gm.GroupID=g.GroupID WHERE g.GroupName=" & Q(grp) & " ORDER BY gm.CommandName", dbOpenSnapshot)
            Do While Not mem.EOF
                AddCommand d, mem!CommandName, rl, grp
                mem.MoveNext
            Loop
            mem.Close
        Else
            AddCommand d, rs!AddrName, rl, ""
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT CommandName, Min(ReceivedOn) AS first_on, Count(*) AS n FROM tblAckReceipts " & _
        "WHERE MessageID=" & msgId & " GROUP BY CommandName", dbOpenSnapshot)
    Do While Not rs.EOF
        nm = rs!CommandName
        If d.Exists(nm) Then
            cur = d(nm)
            d(nm) = Array(cur(0), cur(1), rs!first_on.Value, CLng(rs!n))
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set AddresseeCommands = d
End Function

' Add command cmd to d (AddresseeCommands): TO beats INFO, and each group it
' was reached through is listed once.
Private Sub AddCommand(d As Object, ByVal cmd As String, ByVal rl As String, ByVal grp As String)
    Dim cur As Variant, via As String
    If Not d.Exists(cmd) Then
        d(cmd) = Array(rl, grp, Null, 0&)
        Exit Sub
    End If
    cur = d(cmd)
    If rl = "TO" Then cur(0) = "TO"
    via = cur(1)
    If Len(grp) > 0 And InStr(", " & via & ",", ", " & grp & ",") = 0 Then via = via & IIf(Len(via) > 0, ", ", "") & grp
    cur(1) = via
    d(cmd) = cur
End Sub

' The accountability table's rows, for the page's table: overdue first, then
' in progress, then acknowledged, highest precedence and latest first.
Public Function RptAckSql() As String
    RptAckSql = "SELECT id, ref, subject & IIf(classif In ('UNCLASSIFIED','UNCLASS',''),'','  [' & classif_txt & ']') AS subject_txt, " & _
        "prec, prec_rank, ack_txt, acked, total, status, basis, due, due_txt, received FROM " & ACK_TABLE & _
        " ORDER BY Switch(status='Overdue',1,status='In Progress',2,True,3), prec_rank, received DESC"
End Function

' Totals over ws_rptack: acknowledgements made and owed (people or
' commands), and how many messages are overdue.
' Errors: the engine's.
Public Sub RptTotals(ByRef nAcked As Long, ByRef nOwed As Long, ByRef nOverdue As Long)
    Dim rs As DAO.Recordset
    nAcked = 0: nOwed = 0: nOverdue = 0
    If Not TableExists(ACK_TABLE, True) Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT Sum(acked) AS a, Sum(total) AS t, Sum(IIf(status='Overdue',1,0)) AS o FROM " & _
        ACK_TABLE, dbOpenSnapshot)
    If Not rs.EOF Then
        nAcked = Nz(rs!a, 0): nOwed = Nz(rs!t, 0): nOverdue = Nz(rs!o, 0)
    End If
    rs.Close
End Sub

' One line on message msgId for the detail's heading: "COMSUBLANT 078/26 -
' our people - due 26 Oct 26 18:30 (Overdue by 10 d)".
Public Function RptAckHeading(ByVal msgId As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT ref, basis, due, due_txt, acked, total FROM " & ACK_TABLE & " WHERE id=" & msgId, dbOpenSnapshot)
    If Not rs.EOF Then
        RptAckHeading = rs!ref & " - " & LCase$(rs!basis) & " - " & rs!acked & " of " & rs!total & " acknowledged - "
        ' Not IIf: it would work out AckDueText for a Null due time too.
        If IsNull(rs!due.Value) Then
            RptAckHeading = RptAckHeading & "no due time"
        Else
            RptAckHeading = RptAckHeading & "due " & rs!due_txt & " (" & AckDueText(rs!due.Value) & ")"
        End If
    End If
    rs.Close
End Function

' Fill ws_rptackdet with who must acknowledge message msgId and whether they
' have: the active people for received traffic, the addressee commands for
' ours. Not yet acknowledged shows "Ack due", or "Overdue" past the due time.
' Errors: the engine's.
' Side effects: rewrites ws_rptackdet.
Public Sub RptDetailFill(ByVal msgId As Long)
    Dim db As DAO.Database, rs As DAO.Recordset, basis As String, due As Variant, pending As String, n As Long
    Dim cmds As Object, k As Variant, v As Variant
    Set db = CurrentDb
    EnsureTables
    db.Execute "DELETE FROM " & DET_TABLE, dbFailOnError
    basis = Nz(DLookup("basis", ACK_TABLE, "id=" & msgId), "")
    If Len(basis) = 0 Then Exit Sub
    due = DLookup("AckDue", "tblMessages", "MessageID=" & msgId)
    pending = "Ack due"
    If Not IsNull(due) Then
        If due < Now() Then pending = "Overdue"
    End If
    If basis = BASIS_COMMANDS Then
        Set cmds = AddresseeCommands(msgId)
        For Each k In cmds.Keys
            v = cmds(k)
            n = n + 1
            AddDetail n, msgId, CStr(k), IIf(Len(v(1)) > 0, v(1), "Direct"), CStr(v(0)), v(2), v(3), pending
        Next
    Else
        Set rs = db.OpenRecordset("SELECT p.PersonID, s.AckOn FROM tblPeople AS p LEFT JOIN (SELECT PersonID, AckOn FROM " & _
            "tblMessageUserState WHERE MessageID=" & msgId & ") AS s ON p.PersonID=s.PersonID WHERE p.IsActive", dbOpenSnapshot)
        Do While Not rs.EOF
            n = n + 1
            AddDetail n, msgId, PersonLabel(rs!PersonID), PersonUnit(rs!PersonID), "", rs!AckOn.Value, IIf(IsNull(rs!AckOn.Value), 0, 1), pending
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub

' One ws_rptackdet row. nTimes > 1 adds "(+n more)" to the time: a command
' that acknowledged again (after an amendment).
Private Sub AddDetail(ByVal n As Long, ByVal msgId As Long, ByVal who As String, ByVal via As String, ByVal rl As String, _
        ByVal ackOn As Variant, ByVal nTimes As Long, ByVal pending As String)
    Dim st As String, w As String
    If IsNull(ackOn) Then
        st = pending
    Else
        st = "Acknowledged"
        w = Format$(ackOn, "dd mmm yy hh:nn")
        If nTimes > 1 Then w = w & " (+" & (nTimes - 1) & " more)"
    End If
    CurrentDb.Execute "INSERT INTO " & DET_TABLE & " (id, MessageID, who, via, addr_role, status, ack_on, when_txt) VALUES (" & _
        n & ", " & msgId & ", " & Q(Left$(who, 150)) & ", " & Q(Left$(via, 255)) & ", " & Q(rl) & ", " & Q(st) & ", " & _
        SqlDateOrNull(ackOn) & ", " & Q(w) & ")", dbFailOnError
End Sub

' A person's name from the Personnel file ("LT J. Smith"), or "Person <id>"
' when the file can't be read.
Private Function PersonLabel(ByVal pid As Long) As String
    Dim rs As DAO.Recordset
    On Error GoTo NoName
    Set rs = CurrentDb.OpenRecordset("SELECT rank, first_name, last_name FROM " & PERSON_TABLE & " WHERE person_id=" & pid, dbOpenSnapshot)
    If Not rs.EOF Then PersonLabel = PersonDisplayName(Nz(rs!rank, ""), Nz(rs!first_name, ""), Nz(rs!last_name, ""))
    rs.Close
    If Len(PersonLabel) = 0 Then PersonLabel = "Person " & pid
    Exit Function
NoName:
    PersonLabel = "Person " & pid
End Function

' The detail table's rows: those still owing first.
Public Function RptDetailSql() As String
    RptDetailSql = "SELECT id, who, IIf(addr_role='', via, addr_role & ' - ' & via) AS via_txt, status, when_txt, ack_on FROM " & _
        DET_TABLE & " ORDER BY IIf(status='Acknowledged',1,0), who"
End Function

' Write the accountability table to a CSV file the user picks, and audit
' the export. Returns the path, or "" when the user cancelled.
' Errors: ERR_NOT_ALLOWED without the Report Viewer role; the engine's and
'   the file system's.
' Side effects: writes the file; one audit row.
Public Function RptExportCsv(ByVal periodKey As String) As String
    Dim path As String, n As Long
    RequireRole "REPORT", "export reports"
    path = KitSaveAs("Export the acknowledgement report", "ANCHOR acknowledgements " & Format$(Date, "yyyy-mm-dd") & ".csv", _
        "CSV file|*.csv", "export")
    If Len(path) = 0 Then Exit Function
    If FileExists(path) Then Kill path
    n = ExportSqlToCsv("SELECT ref AS Reference, subject AS [Message subject], prec AS Precedence, classif_txt AS Classification, " & _
        "basis AS [Who acknowledges], due_txt AS [Ack due], acked AS Acknowledged, total AS [Expected], status AS [Ack status] FROM " & _
        ACK_TABLE & " ORDER BY Switch(status='Overdue',1,status='In Progress',2,True,3), prec_rank, received DESC", path)
    AuditLog "Export", "Acknowledgement report", "Acknowledgement accountability exported to CSV (" & n & " " & _
        Plural(n, "message") & "; " & RptPeriodText(periodKey) & ")."
    KitToast "Exported " & n & " " & Plural(n, "row") & " to " & Mid$(path, InStrRev(path, "\") + 1), "success"
    RptExportCsv = path
End Function

' ---------------------------------------------------------------------------
' Traffic measures
' ---------------------------------------------------------------------------

' Counts of the period's traffic by a measure: "prec" precedence, "type"
' message type (named from tblMessageTypes), "class" classification (with
' the demo marking note for a level above UNCLASSIFIED) or "life"
' lifecycle. Columns: id and label (the value), n, share_txt, note_txt.
' total is the period's message count (RptMessageCount), at least 1.
Public Function RptTrafficSql(ByVal measure As String, ByVal periodKey As String, ByVal total As Long) As String
    Dim col As String, noteSql As String, ord As String, src As String
    If total < 1 Then total = 1
    src = "tblMessages AS m"
    Select Case measure
        Case "type"
            col = "Nz(m.MsgType,'(none)')"
            src = "tblMessages AS m LEFT JOIN tblMessageTypes AS t ON m.MsgType=t.TypeCode"
            noteSql = "Nz(First(t.TypeName),'')"
            ord = "Count(*) DESC, " & col
        Case "class"
            col = "Nz(m.Classification,'UNCLASSIFIED')"
            noteSql = "IIf(" & col & " In ('UNCLASSIFIED','UNCLASS',''),''," & Q(DEMO_MARKING_NOTE) & ")"
            ord = "Switch(" & col & "='SECRET',1," & col & "='CONFIDENTIAL',2," & col & "='CUI',3,True,4)"
        Case "life"
            col = "Nz(m.Lifecycle,'(none)')"
            noteSql = "''"
            ord = "Count(*) DESC, " & col
        Case Else
            col = "Nz(m.Precedence,'(none)')"
            noteSql = "''"
            ord = "Min(" & PrecedenceRankSql() & ")"
    End Select
    RptTrafficSql = "SELECT " & col & " AS id, " & col & " AS label, Count(*) AS n, Format(Count(*)/" & total & ",'0%') AS share_txt, " & _
        noteSql & " AS note_txt FROM " & src & " WHERE " & RptPeriodSql(periodKey) & " GROUP BY " & col & " ORDER BY " & ord
End Function
