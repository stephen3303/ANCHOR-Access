Option Compare Database
Option Explicit

' Messages (see Build_pgMessages). Every message newest first
' (modMessageView.MessagesListSql over modMessages' shared list columns):
' the unread dot, precedence, lifecycle and Ack due. The search box looks in
' the subject, reference and originator; All / Unread / Ack due and the
' precedence, type (tblMessageTypes) and classification combos narrow it.
' The picked message's preview shows on the right: subject, reference, DTG,
' originator, its pills, the demo marking note above UNCLASSIFIED, the
' lifecycle notice and the first lines of the official text. Open (or a
' double-click) opens it on the Message page (OpenMessage), which marks it
' read. Mark read and Mark unread change the picked one, Mark all read every
' message; they change only the signed-in person's own read state.
'
' Depends on: modGlobals, modMessages, modMessageView, modSignIn, modUI;
' the kit's modKitUI, modKitUtil, modMenu.

Private Const PREVIEW_CHARS As Long = 900       ' how much of the official text the preview shows

Private m_search As String
Private m_id As Long            ' the message shown on the right (0: none)

Public Property Get HelpTopic() As String
    HelpTopic = "messages"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    ResetCombos
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no messages."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        SetEnabled Me.btnMarkRead, False
        SetEnabled Me.btnMarkUnread, False
        SetEnabled Me.btnMarkAll, False
        Me.lblSubtitle.Caption = ""
        ShowPreview 0
        Exit Sub
    End If
    RefreshList 0
    Exit Sub
Err_Handler:
    HandleError "pgMessages.Form_Load"
End Sub

' Each combo on its first entry, which means no filter; the types from
' tblMessageTypes.
Private Sub ResetCombos()
    Me.cboType.RowSourceType = "Value List"
    Me.cboType.RowSource = MessageTypeValueList()
    Me.cboPrec = MSG_ALL_PREC
    Me.cboType = MSG_ALL_TYPES
    Me.cboClass = MSG_ALL_CLASSES
End Sub

' ---------------------------------------------------------------------------
' Filters and search (the kit's builders call these)
' ---------------------------------------------------------------------------

' All / Unread / Ack due changed (segShow's onChange).
Public Sub ShowChanged()
    On Error GoTo Err_Handler
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgMessages.ShowChanged"
End Sub

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList m_id
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgMessages.SearchFor", text
End Sub

Private Sub cboPrec_AfterUpdate()
    On Error GoTo Err_Handler
    If IsNull(Me.cboPrec) Then Me.cboPrec = MSG_ALL_PREC
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgMessages.cboPrec_AfterUpdate"
End Sub

Private Sub cboType_AfterUpdate()
    On Error GoTo Err_Handler
    If IsNull(Me.cboType) Then Me.cboType = MSG_ALL_TYPES
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgMessages.cboType_AfterUpdate"
End Sub

Private Sub cboClass_AfterUpdate()
    On Error GoTo Err_Handler
    If IsNull(Me.cboClass) Then Me.cboClass = MSG_ALL_CLASSES
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgMessages.cboClass_AfterUpdate"
End Sub

' The empty state's button: every filter and the search cleared.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    SetSegment Me, "segShow", MSG_SHOW_ALL
    Me.cboPrec = MSG_ALL_PREC
    Me.cboType = MSG_ALL_TYPES
    Me.cboClass = MSG_ALL_CLASSES
    SetSearch Me, "txtSearch", ""            ' calls SearchFor, which loads the list
    Exit Sub
Err_Handler:
    HandleError "pgMessages.emptyList_b_Click"
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Private Sub btnOpen_Click()
    On Error GoTo Err_Handler
    If m_id <> 0 Then OpenMessage m_id
    Exit Sub
Err_Handler:
    HandleError "pgMessages.btnOpen_Click", m_id
End Sub

Private Sub btnMarkRead_Click()
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    MarkRead m_id
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgMessages.btnMarkRead_Click", m_id
End Sub

Private Sub btnMarkUnread_Click()
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    MarkUnread m_id
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgMessages.btnMarkUnread_Click", m_id
End Sub

Private Sub btnMarkAll_Click()
    Dim n As Long
    On Error GoTo Err_Handler
    n = MarkAllRead()
    RefreshList m_id
    If n = 0 Then
        KitToast "Every message was already read."
    Else
        KitToast "Marked " & n & " " & Plural(n, "message") & " read.", "success"
    End If
    Exit Sub
Err_Handler:
    HandleError "pgMessages.btnMarkAll_Click"
End Sub

' A row clicked: its preview on the right; a double-click opens it.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    If grid <> "sfList" Then Exit Sub
    id = Nz(Me.sfList.Form.txtSel, 0)
    ShowPreview id
    If dbl And id <> 0 Then OpenMessage id
    Exit Sub
Err_Handler:
    HandleError "pgMessages.GridClick", grid & " " & id
End Sub

' A right-click on a row: its actions.
Public Sub GridRightClick(ByVal grid As String)
    Dim k As String, isRead As Boolean
    On Error GoTo Err_Handler
    If grid <> "sfList" Or m_id = 0 Then Exit Sub
    isRead = (ReadStateOf(m_id) = "read")
    k = KitMenuAt("open|Open|E8A7", IIf(isRead, "~", "") & "read|Mark read|E8C3", IIf(isRead, "", "~") & "unread|Mark unread|E715")
    Select Case k
        Case "open": OpenMessage m_id
        Case "read": btnMarkRead_Click
        Case "unread": btnMarkUnread_Click
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgMessages.GridRightClick", grid
End Sub

' ---------------------------------------------------------------------------
' The list and the preview
' ---------------------------------------------------------------------------

' Load the messages for the filters and the search, and show openId (0, or
' not listed: the first).
Private Sub RefreshList(ByVal openId As Long)
    Dim n As Long, unread As Long, due As Long, total As Long
    If Not DataLinked() Then Exit Sub
    Me.sfList.Form.RecordSource = MessagesListSql(SegmentValue(Me, "segShow"), Nz(Me.cboPrec, ""), Nz(Me.cboType, ""), _
        Nz(Me.cboClass, ""), m_search)
    GridHighlight Me.sfList, m_search
    n = Me.sfList.Form.RecordsetClone.RecordCount
    total = DCount("*", "tblMessages")
    unread = UnreadCount()
    due = PendingAckCount()
    Me.lblSubtitle.Caption = IIf(n = total, total & " " & Plural(total, "message"), n & " of " & total & " messages") & _
        ", " & unread & " unread, " & due & " " & Plural(due, "acknowledgement") & " due from you"
    If total = 0 Then
        ShowEmpty Me, "emptyList", True, "No messages yet", "Messages ANCHOR receives from C2OIX show here, newest first."
        SetVisible Me.emptyList_b, False
    Else
        ShowEmpty Me, "emptyList", n = 0, "No messages match", "No message fits the search and the filters."
    End If
    SetVisible Me.sfList, n > 0
    SetEnabled Me.btnMarkAll, unread > 0
    If GridSelect(Me.sfList, openId) Then
        ShowPreview openId
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then      ' a row found calls GridClick, which shows it
        ShowPreview 0
    End If
End Sub

' Show message msgId's preview on the right (0: the hint instead).
Private Sub ShowPreview(ByVal msgId As Long)
    Dim rs As DAO.Recordset, isShown As Boolean, c As Control, noteTxt As String, body As String, isRead As Boolean
    m_id = 0
    If msgId <> 0 And DataLinked() Then
        Set rs = CurrentDb.OpenRecordset("SELECT Subject, Reference, Dtg, Originator, Precedence, Classification, Lifecycle, " & _
            "ReplacedBy, RelatesTo, AckDue, BodyText FROM tblMessages WHERE MessageID=" & msgId, dbOpenSnapshot)
        If Not rs.EOF Then
            isShown = True
            m_id = msgId
            Me.lblPvSubject.Caption = Replace(Nz(rs!Subject, ""), "&", "&&")
            Me.lblPvRef.Caption = Replace(Nz(rs!Reference, "") & " | " & Nz(rs!Dtg, ""), "&", "&&")
            Me.lblPvFrom.Caption = Replace("From " & Nz(rs!Originator, ""), "&", "&&")
            SetPill Me.pvPrec, rs!Precedence
            SetPill Me.pvClass, rs!Classification
            SetPill Me.pvLife, rs!Lifecycle
            If IsMarkedAboveUnclass(rs!Classification) Then noteTxt = ClassificationMarking(rs!Classification)
            AddLine noteTxt, LifecycleNotice(Nz(rs!Lifecycle, ""), rs!ReplacedBy, rs!RelatesTo)
            If NeedsOwnAck(msgId) Then AddLine noteTxt, "Your acknowledgement is required. " & AckDueText(rs!AckDue)
            Me.lblPvNote.Caption = Replace(noteTxt, "&", "&&")
            body = ShowLines(rs!BodyText)
            If Len(body) > PREVIEW_CHARS Then body = RTrim$(Left$(body, PREVIEW_CHARS)) & ChrW(8230)
            Me.txtPvBody = body
            isRead = (ReadStateOf(msgId) = "read")
        End If
        rs.Close
    End If
    For Each c In Me.Controls
        Select Case c.Name
            Case "lblPvSubject", "lblPvRef", "lblPvFrom", "pvPrec", "pvClass", "pvLife", "lblPvNote", "capPvBody", "txtPvBody", "btnOpen"
                SetVisible c, isShown
        End Select
    Next
    SetVisible Me.lblDetailEmpty, Not isShown
    SetEnabled Me.btnMarkRead, isShown And Not isRead
    SetEnabled Me.btnMarkUnread, isShown And isRead
End Sub

' Put a word in a pill (PillBox) and repaint its rounded image.
Private Sub SetPill(ByVal c As Control, ByVal v As Variant)
    c.Value = Nz(v, "")
    Me.Controls(c.Name & "_bg").Requery
End Sub

' Add line s to text t ("" adds nothing).
Private Sub AddLine(t As String, ByVal s As String)
    If Len(s) = 0 Then Exit Sub
    t = t & IIf(Len(t) > 0, vbCrLf, "") & s
End Sub
