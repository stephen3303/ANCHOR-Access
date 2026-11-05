Option Compare Database
Option Explicit

' Watch turnover (see modBuildWatch.Build_pgWatchTurnover). The tiles are
' the web app's watch prep counts for the signed-in person: unread messages
' and acknowledgements they still owe (modMessages.UnreadCount,
' PendingAckCount), the open action items in current traffic
' (modWatch.WatchOpenActionCount), and the open turnover's checklist
' progress.
'
' Watch notes: open ones first; Add note asks for the text, Close closes the
' picked open note. Ongoing topics: Add topic asks for its name (it starts
' Active), Status sets the picked topic Active, Monitoring or Closed (and
' its last update to now). The checklist: Start turnover opens a turnover;
' while it is open a click on an item ticks or unticks it; Complete
' turnover needs every item ticked. Everyone may use the page; each write
' is checked and audited in modWatch.
'
' Depends on: modMessages, modSignIn, modUI, modWatch; the kit's modKitUI,
' modMenu.

Private Const MSG_TITLE As String = "Watch turnover"

Private m_turnId As Long        ' the open turnover (0: none)

Public Property Get HelpTopic() As String
    HelpTopic = "watch"
End Property

Private Sub Form_Load()
    Dim c As Variant
    On Error GoTo Err_Handler
    If Not DataLinked() Then
        For Each c In Array("tileUnread", "tileAcks", "tileActions", "tileChecks")
            Me.Controls(CStr(c)).Value = Null
        Next
        For Each c In Array("btnAddNote", "btnCloseNote", "btnAddTopic", "btnStatus", "btnStart", "btnComplete")
            SetEnabled Me.Controls(CStr(c)), False
        Next
        Me.lblTurn.Caption = "This copy was built without the data file, so there is no watch to turn over."
        ShowEmpty Me, "emptyNotes", True, "No data file linked", "There are no watch notes to show."
        SetVisible Me.sfNotes, False
        ShowEmpty Me, "emptyTopics", True, "No data file linked", "There are no topics to show."
        SetVisible Me.sfTopics, False
        ShowEmpty Me, "emptyChecks", True, "No data file linked", "There is no checklist to show."
        SetVisible Me.sfChecks, False
        Exit Sub
    End If
    Me.sfNotes.Form.RecordSource = WatchNotesSql()
    Me.sfTopics.Form.RecordSource = WatchTopicsSql()
    ShowNotesEmpty
    ShowTopicsEmpty
    SelectFirstRow Me.sfNotes
    SelectFirstRow Me.sfTopics
    RefreshCounts
    RefreshChecks
    Exit Sub
Err_Handler:
    HandleError "pgWatchTurnover.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' Tiles and empty states
' ---------------------------------------------------------------------------

Private Sub RefreshCounts()
    Me.tileUnread = UnreadCount()
    Me.tileAcks = PendingAckCount()
    Me.tileActions = WatchOpenActionCount()
End Sub

Private Sub ShowNotesEmpty()
    Dim n As Long
    n = Me.sfNotes.Form.RecordsetClone.RecordCount
    ShowEmpty Me, "emptyNotes", n = 0
    SetVisible Me.sfNotes, n > 0
    SetEnabled Me.btnCloseNote, n > 0
End Sub

Private Sub ShowTopicsEmpty()
    Dim n As Long
    n = Me.sfTopics.Form.RecordsetClone.RecordCount
    ShowEmpty Me, "emptyTopics", n = 0
    SetVisible Me.sfTopics, n > 0
    SetEnabled Me.btnStatus, n > 0
End Sub

' ---------------------------------------------------------------------------
' Watch notes
' ---------------------------------------------------------------------------

Private Sub btnAddNote_Click()
    Dim v As Variant, id As Long
    On Error GoTo Err_Handler
    v = KitInputBox("The note for the watch:", "Add a watch note")
    If IsNull(v) Then Exit Sub
    If Len(Trim$(v)) = 0 Then
        KitMsgBox "The note was empty, so nothing was added.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    id = WatchAddNote(CStr(v))
    GridSelect Me.sfNotes, id, True
    ShowNotesEmpty
    KitToast "Watch note added.", "success"
    Exit Sub
Err_Handler:
    HandleError "pgWatchTurnover.btnAddNote_Click"
End Sub

Private Sub btnCloseNote_Click()
    Dim id As Long
    On Error GoTo Err_Handler
    id = Nz(Me.sfNotes.Form.txtSel, 0)
    If id = 0 Then
        KitMsgBox "Pick the note to close first.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    If WatchCloseNote(id) Then
        KitToast "Watch note closed.", "success"
    Else
        KitMsgBox "That note is already closed.", vbInformation, MSG_TITLE
    End If
    GridSelect Me.sfNotes, id, True
    ShowNotesEmpty
    Exit Sub
Err_Handler:
    HandleError "pgWatchTurnover.btnCloseNote_Click"
End Sub

' ---------------------------------------------------------------------------
' Ongoing topics
' ---------------------------------------------------------------------------

Private Sub btnAddTopic_Click()
    Dim v As Variant, id As Long, nm As String
    On Error GoTo Err_Handler
    v = KitInputBox("The topic's name:", "Add an ongoing topic")
    If IsNull(v) Then Exit Sub
    nm = Trim$(v)
    If Len(nm) = 0 Then Exit Sub
    If DCount("*", "tblOngoingTopics", "Topic=" & Q(nm)) > 0 Then
        KitMsgBox "The topic " & nm & " is already listed. Pick it and change its status instead.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    id = WatchAddTopic(nm)
    GridSelect Me.sfTopics, id, True
    ShowTopicsEmpty
    KitToast "Topic added as Active.", "success"
    Exit Sub
Err_Handler:
    HandleError "pgWatchTurnover.btnAddTopic_Click"
End Sub

Private Sub btnStatus_Click()
    Dim id As Long, pick As String
    On Error GoTo Err_Handler
    id = Nz(Me.sfTopics.Form.txtSel, 0)
    If id = 0 Then
        KitMsgBox "Pick the topic first.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    pick = KitMenu(Me.btnStatus, "Active|Active|E768", "Monitoring|Monitoring|E7B3", "Closed|Closed|E711")
    If Len(pick) = 0 Then Exit Sub
    If WatchSetTopicStatus(id, pick) Then KitToast "Topic set " & pick & ".", "success"
    GridSelect Me.sfTopics, id, True
    Exit Sub
Err_Handler:
    HandleError "pgWatchTurnover.btnStatus_Click"
End Sub

' ---------------------------------------------------------------------------
' The checklist
' ---------------------------------------------------------------------------

' The open turnover's line, its checklist and progress, and which of Start
' and Complete can be used.
Private Sub RefreshChecks()
    Dim nDone As Long, nItems As Long, keep As Variant
    m_turnId = WatchOpenTurnover()
    keep = Me.sfChecks.Form.txtSel
    Me.sfChecks.Form.RecordSource = WatchChecksSql(m_turnId)
    WatchCheckCounts m_turnId, nDone, nItems
    Me.lblTurn.Caption = Replace(WatchTurnoverText(m_turnId), "&", "&&")
    If m_turnId = 0 Then
        Me.tileChecks = "Not started"
        Me.txtPct = 0
    Else
        Me.tileChecks = nDone & " of " & nItems
        If nItems > 0 Then Me.txtPct = CLng(100 * nDone / nItems) Else Me.txtPct = 0
    End If
    ShowEmpty Me, "emptyChecks", nItems = 0
    SetVisible Me.sfChecks, nItems > 0
    SetEnabled Me.btnStart, m_turnId = 0
    SetEnabled Me.btnComplete, m_turnId <> 0
    GridSelect Me.sfChecks, keep
    Me.Recalc                               ' the progress bar reads txtPct
End Sub

' A row clicked: a checklist item is ticked or unticked for the open
' turnover.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim itemId As Long
    On Error GoTo Err_Handler
    If grid <> "sfChecks" Or dbl Then Exit Sub      ' a double click's second click would untick it again
    itemId = Nz(Me.sfChecks.Form.txtSel, 0)
    If itemId = 0 Then Exit Sub
    If m_turnId = 0 Then
        KitMsgBox "Start a turnover first: the checklist is ticked for one turnover at a time.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    WatchToggleCheck m_turnId, itemId
    RefreshChecks
    Exit Sub
Err_Handler:
    HandleError "pgWatchTurnover.GridClick", grid
End Sub

Private Sub btnStart_Click()
    On Error GoTo Err_Handler
    If WatchOpenTurnover() <> 0 Then
        KitMsgBox "Someone started a turnover meanwhile. It shows now.", vbInformation, MSG_TITLE
    Else
        WatchStartTurnover
        KitToast "Turnover started. Tick each item as you finish it.", "success"
    End If
    RefreshChecks
    RefreshCounts
    Exit Sub
Err_Handler:
    HandleError "pgWatchTurnover.btnStart_Click"
End Sub

Private Sub btnComplete_Click()
    Dim nDone As Long, nItems As Long
    On Error GoTo Err_Handler
    If m_turnId = 0 Then Exit Sub
    WatchCheckCounts m_turnId, nDone, nItems
    If nDone < nItems Then
        KitMsgBox "Tick every item on the checklist first: " & (nItems - nDone) & " " & IIf(nItems - nDone = 1, "is", "are") & _
            " not ticked yet.", vbExclamation, MSG_TITLE
        Exit Sub
    End If
    If KitMsgBox("Complete the turnover? You take the watch, and it is recorded in the audit trail.", vbOKCancel + vbQuestion, _
            MSG_TITLE) <> vbOK Then Exit Sub
    WatchCompleteTurnover m_turnId
    KitToast "Turnover completed. You have the watch.", "success"
    RefreshChecks
    RefreshCounts
    Exit Sub
Err_Handler:
    HandleError "pgWatchTurnover.btnComplete_Click"
End Sub
