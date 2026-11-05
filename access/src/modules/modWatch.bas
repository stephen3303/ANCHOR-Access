Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the watch turnover: the watch notes
' (tblWatchNotes, open until someone closes them), the ongoing topics
' (tblOngoingTopics: Active, Monitoring or Closed, with when each last
' changed), and the turnover checklist.
'
' A turnover is one row of tblTurnovers, from Start turnover until Complete
' turnover; while it is open the checklist's items (tblTurnoverChecklist,
' the active ones) are ticked for it in tblTurnoverChecks. Only one turnover
' is open at a time: the current one is the open row. Complete needs every
' active item ticked. Starting and completing a turnover, notes and topics
' are audited; a single tick isn't (the completion records that all were).
'
' Everyone who may use ANCHOR may use the page: each write checks the
' Message Viewer role, which every ANCHOR user holds, in the data file
' (RequireRole), so a person turned off since signing in can't write.
'
' Owner: the Watch turnover page. Depends on: modAudit, modErrors,
' modGlobals, modSignIn; the kit's modKitLocks.

' The ongoing topics' statuses, in their order.
Public Const WATCH_TOPIC_STATUSES As String = "Active;Monitoring;Closed"

' ---------------------------------------------------------------------------
' Counts for the tiles
' ---------------------------------------------------------------------------

' How many messages carry an action item (tblMessages.ActionText) and still
' direct action: not superseded, cancelled or expired.
' Errors: the engine's.
Public Function WatchOpenActionCount() As Long
    WatchOpenActionCount = DCount("*", "tblMessages", "Len(Nz(ActionText,''))>0 AND Lifecycle Not In ('SUPERSEDED','CANCELLED','EXPIRED')")
End Function

' ---------------------------------------------------------------------------
' Watch notes
' ---------------------------------------------------------------------------

' The watch notes, open ones first, newest first.
Public Function WatchNotesSql() As String
    WatchNotesSql = "SELECT n.NoteID AS id, IIf(n.IsOpen,'Open','Closed') AS note_state, Format(n.CreatedOn,'dd mmm hh:nn') AS created_txt, " & _
        "n.CreatedOn, Nz(n.PersonName,'') AS who, n.NoteText AS note_text FROM tblWatchNotes AS n ORDER BY n.IsOpen, n.CreatedOn DESC, n.NoteID DESC"
End Function

' Add an open watch note by the signed-in person. Returns its id.
' Errors: ERR_NOT_ALLOWED; 5 for an empty note; the engine's.
' Side effects: one audit row.
Public Function WatchAddNote(ByVal noteText As String) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, id As Long
    noteText = Trim$(noteText)
    If Len(noteText) = 0 Then Err.Raise 5, "modWatch.WatchAddNote", "Write the note first."
    RequireRole "VIEW", "add watch notes"
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    id = InsertRow("INSERT INTO tblWatchNotes (NoteText, PersonID, PersonName, CreatedOn, IsOpen) VALUES (" & Q(noteText) & ", " & _
        SignedInPersonID() & ", " & Q(Left$(SignedInName(), 100)) & ", Now(), True)")
    AuditLog "Add watch note", "Watch notes", "Watch note added: " & Left$(noteText, 200)
    ws.CommitTrans: inTrans = False
    WatchAddNote = id
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modWatch.WatchAddNote"
End Function

' Close watch note noteId. Returns False when it was already closed (or is
' gone), which changes nothing.
' Errors: ERR_NOT_ALLOWED; the engine's.
' Side effects: one audit row when it closes.
Public Function WatchCloseNote(ByVal noteId As Long) As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, txt As String
    RequireRole "VIEW", "close watch notes"
    txt = Nz(DLookup("NoteText", "tblWatchNotes", "NoteID=" & noteId), "")
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If KitExecute("UPDATE tblWatchNotes SET IsOpen=False WHERE NoteID=" & noteId & " AND IsOpen") > 0 Then
        AuditLog "Close watch note", "Watch notes", "Watch note closed: " & Left$(txt, 200)
        WatchCloseNote = True
    End If
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modWatch.WatchCloseNote", noteId
End Function

' ---------------------------------------------------------------------------
' Ongoing topics
' ---------------------------------------------------------------------------

' The ongoing topics: Active, then Monitoring, then Closed, by name.
Public Function WatchTopicsSql() As String
    WatchTopicsSql = "SELECT t.TopicID AS id, t.Topic AS topic_name, Nz(t.Status,'Active') AS topic_status, " & _
        "Format(t.LastUpdate,'dd mmm yyyy hh:nn') AS upd_txt, t.LastUpdate FROM tblOngoingTopics AS t " & _
        "ORDER BY Switch(t.Status='Active',1,t.Status='Monitoring',2,True,3), t.Topic"
End Function

' Add an Active topic, updated now. Returns its id.
' Errors: ERR_NOT_ALLOWED; 5 for an empty name or one already listed; the
'   engine's.
' Side effects: one audit row.
Public Function WatchAddTopic(ByVal topicName As String) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, id As Long
    topicName = Left$(Trim$(topicName), 100)
    If Len(topicName) = 0 Then Err.Raise 5, "modWatch.WatchAddTopic", "Name the topic first."
    If DCount("*", "tblOngoingTopics", "Topic=" & Q(topicName)) > 0 Then
        Err.Raise 5, "modWatch.WatchAddTopic", "The topic " & topicName & " is already listed. Change its status instead."
    End If
    RequireRole "VIEW", "add ongoing topics"
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    id = InsertRow("INSERT INTO tblOngoingTopics (Topic, LastUpdate, Status) VALUES (" & Q(topicName) & ", Now(), 'Active')")
    AuditLog "Add topic", "Ongoing topic (" & topicName & ")", "Ongoing topic added as Active."
    ws.CommitTrans: inTrans = False
    WatchAddTopic = id
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modWatch.WatchAddTopic", topicName
End Function

' Set topic topicId's status (one of WATCH_TOPIC_STATUSES) and its last
' update to now. Returns False when it already had that status.
' Errors: ERR_NOT_ALLOWED; 5 for another status or a topic that is gone;
'   the engine's.
' Side effects: one audit row when it changes.
Public Function WatchSetTopicStatus(ByVal topicId As Long, ByVal newStatus As String) As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, cur As Variant, nm As String
    If InStr(";" & WATCH_TOPIC_STATUSES & ";", ";" & newStatus & ";") = 0 Or Len(newStatus) = 0 Then
        Err.Raise 5, "modWatch.WatchSetTopicStatus", "A topic is Active, Monitoring or Closed, not " & newStatus & "."
    End If
    RequireRole "VIEW", "change ongoing topics"
    nm = Nz(DLookup("Topic", "tblOngoingTopics", "TopicID=" & topicId), "")
    If Len(nm) = 0 Then Err.Raise 5, "modWatch.WatchSetTopicStatus", "That topic was deleted."
    cur = DLookup("Status", "tblOngoingTopics", "TopicID=" & topicId)
    If StrComp(Nz(cur, ""), newStatus, vbBinaryCompare) = 0 Then Exit Function
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "UPDATE tblOngoingTopics SET Status=" & Q(newStatus) & ", LastUpdate=Now() WHERE TopicID=" & topicId
    AuditLog "Change topic", "Ongoing topic (" & nm & ")", "Status changed from " & Nz(cur, "none") & " to " & newStatus & "."
    ws.CommitTrans: inTrans = False
    WatchSetTopicStatus = True
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modWatch.WatchSetTopicStatus", topicId & " " & newStatus
End Function

' ---------------------------------------------------------------------------
' Turnovers and the checklist
' ---------------------------------------------------------------------------

' The open turnover's id, 0 when none is open.
Public Function WatchOpenTurnover() As Long
    WatchOpenTurnover = Nz(DMax("TurnoverID", "tblTurnovers", "CompletedOn Is Null"), 0)
End Function

' One line on turnover turnId: "Started 05 Nov 2026 14:00 by LT J. Smith."
' or, with none open, when the last one was completed.
Public Function WatchTurnoverText(ByVal turnId As Long) As String
    Dim rs As DAO.Recordset, lastId As Long
    If turnId = 0 Then
        lastId = Nz(DMax("TurnoverID", "tblTurnovers", "CompletedOn Is Not Null"), 0)
        If lastId = 0 Then
            WatchTurnoverText = "No turnover started. Start one when you take over the watch."
        Else
            Set rs = CurrentDb.OpenRecordset("SELECT PersonName, CompletedOn FROM tblTurnovers WHERE TurnoverID=" & lastId, dbOpenSnapshot)
            WatchTurnoverText = "No turnover open. The last one was completed " & Format$(rs!CompletedOn, "dd mmm yyyy hh:nn") & _
                " by " & Nz(rs!PersonName, "someone") & "."
            rs.Close
        End If
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT PersonName, StartedOn FROM tblTurnovers WHERE TurnoverID=" & turnId, dbOpenSnapshot)
    If Not rs.EOF Then
        WatchTurnoverText = "Turnover started " & Format$(rs!StartedOn, "dd mmm yyyy hh:nn") & " by " & Nz(rs!PersonName, "someone") & _
            ". Tick each item as you finish it."
    End If
    rs.Close
End Function

' Start a turnover by the signed-in person. Returns its id.
' Errors: ERR_NOT_ALLOWED; 5 when one is already open; the engine's.
' Side effects: one audit row.
Public Function WatchStartTurnover() As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, id As Long
    RequireRole "VIEW", "start a turnover"
    If WatchOpenTurnover() <> 0 Then Err.Raise 5, "modWatch.WatchStartTurnover", "A turnover is already open. Complete it first."
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    id = InsertRow("INSERT INTO tblTurnovers (PersonID, PersonName, StartedOn) VALUES (" & SignedInPersonID() & ", " & _
        Q(Left$(SignedInName(), 100)) & ", Now())")
    AuditLog "Start turnover", "Watch turnover", "Watch turnover started."
    ws.CommitTrans: inTrans = False
    WatchStartTurnover = id
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modWatch.WatchStartTurnover"
End Function

' The checklist's active items for turnover turnId (0: none open, every
' item Not Started): id (ItemID), check_state Done or Not Started, the
' item, when it was ticked.
Public Function WatchChecksSql(ByVal turnId As Long) As String
    WatchChecksSql = "SELECT c.ItemID AS id, IIf(k.CheckedOn Is Null,'Not Started','Done') AS check_state, c.ItemText AS item_text, " & _
        "Format(k.CheckedOn,'hh:nn') AS checked_txt, c.SortOrder FROM tblTurnoverChecklist AS c LEFT JOIN (SELECT ItemID, CheckedOn " & _
        "FROM tblTurnoverChecks WHERE TurnoverID=" & turnId & ") AS k ON c.ItemID=k.ItemID WHERE c.IsActive ORDER BY c.SortOrder, c.ItemID"
End Function

' How many of the active checklist items are ticked for turnover turnId,
' and how many there are.
' Errors: the engine's.
Public Sub WatchCheckCounts(ByVal turnId As Long, ByRef nDone As Long, ByRef nItems As Long)
    Dim rs As DAO.Recordset
    nItems = DCount("*", "tblTurnoverChecklist", "IsActive")
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM tblTurnoverChecks AS k INNER JOIN tblTurnoverChecklist AS c " & _
        "ON k.ItemID=c.ItemID WHERE c.IsActive AND k.TurnoverID=" & turnId, dbOpenSnapshot)
    nDone = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Sub

' Tick checklist item itemId for open turnover turnId, or untick it.
' Returns True when it is ticked now.
' Errors: ERR_NOT_ALLOWED; 5 when the turnover isn't open any more; the
'   engine's.
Public Function WatchToggleCheck(ByVal turnId As Long, ByVal itemId As Long) As Boolean
    RequireRole "VIEW", "tick the turnover checklist"
    If DCount("*", "tblTurnovers", "TurnoverID=" & turnId & " AND CompletedOn Is Null") = 0 Then
        Err.Raise 5, "modWatch.WatchToggleCheck", "This turnover was completed. Start a new one to tick the checklist again."
    End If
    If DCount("*", "tblTurnoverChecks", "TurnoverID=" & turnId & " AND ItemID=" & itemId) > 0 Then
        KitExecute "DELETE FROM tblTurnoverChecks WHERE TurnoverID=" & turnId & " AND ItemID=" & itemId
    Else
        KitExecute "INSERT INTO tblTurnoverChecks (TurnoverID, ItemID, CheckedOn) VALUES (" & turnId & ", " & itemId & ", Now())"
        WatchToggleCheck = True
    End If
End Function

' Complete open turnover turnId: every active checklist item must be ticked.
' Errors: ERR_NOT_ALLOWED; 5 when an item isn't ticked or the turnover was
'   completed meanwhile; the engine's and KitExecute's.
' Side effects: one audit row, in the same transaction.
Public Sub WatchCompleteTurnover(ByVal turnId As Long)
    Dim ws As DAO.Workspace, inTrans As Boolean, nDone As Long, nItems As Long
    RequireRole "VIEW", "complete a turnover"
    WatchCheckCounts turnId, nDone, nItems
    If nDone < nItems Then
        Err.Raise 5, "modWatch.WatchCompleteTurnover", "Tick every item on the checklist first: " & (nItems - nDone) & " " & _
            IIf(nItems - nDone = 1, "is", "are") & " not ticked yet."
    End If
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If KitExecute("UPDATE tblTurnovers SET CompletedOn=Now() WHERE TurnoverID=" & turnId & " AND CompletedOn Is Null") = 0 Then
        Err.Raise 5, "modWatch.WatchCompleteTurnover", "This turnover was already completed."
    End If
    AuditLog "Complete turnover", "Watch turnover", "Watch turnover completed with all " & nItems & " checklist " & _
        Plural(nItems, "item") & " ticked."
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modWatch.WatchCompleteTurnover", turnId
End Sub
