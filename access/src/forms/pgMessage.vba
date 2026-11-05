Option Compare Database
Option Explicit

' Message (see Build_pgMessage). The message TempVars OpenMessageID names
' (modMessages.OpenMessage sets it from any list), marked read on opening
' (MarkRead, which audits the first view). The subject, reference,
' originator and DTG, its classification, precedence and lifecycle pills
' with the demo marking note above UNCLASSIFIED; the lifecycle banner
' (LifecycleNotice) with a link to the replacing or related message when
' ANCHOR holds it; the acknowledgement bar: Acknowledge while the person
' must still acknowledge it, "Acknowledged <time>" once they have, and for
' this command's own traffic how many addressee commands acknowledged.
' The official text (BodyText) in a large read-only box, switched to the
' raw text exactly as received (RawText, monospace); the stored AI summary
' apart in its own card, labelled not authoritative. On the right, four
' tabs: Details (type, class, the integrity data recorded at ingestion,
' the action item), Addressees (with the distribution and the attachments'
' names, no files), Acks (own traffic: each addressee command, groups
' expanded through tblGroupMembers, against tblAckReceipts, as the web
' app's resolveAckStatus) and Related (a click opens one here). Print opens
' rptMessage, the official rendition.
'
' A related message opens on this page without leaving it: the click sets
' the TempVar and the page loads it from its timer, once the click's own
' event has finished. The first message loads the same way, once the page
' is open and sized (Form_Load only lists the controls and sets the
' timer), so no subform is read or requeried, and nothing is written
' (MarkRead), while the shell is still loading the page.
'
' Depends on: modAudit, modGlobals, modMessages, modMessageView, modSignIn,
' modUI; the kit's modKitUI, modKitUtil.

Private Const RPT_MESSAGE As String = "rptMessage"
Private Const ERR_REPORT_CANCELLED As Long = 2501   ' OpenReport when the report cancelled its own opening

' The controls shown while a message is, apart from the header (Form_Load
' lists them, ";a;b;"), so a missing message hides them all.
Private m_content As String
Private m_id As Long            ' the message shown (0: none)
Private m_tab As String         ' the open tab: d, a, k or r
Private m_lifeRef As String     ' the message lnkLife opens
Private m_nextId As Long        ' the message Form_Timer loads next
Private m_attN As Long          ' attachments of the message shown
Private m_relN As Long          ' related messages ANCHOR holds
Private m_ownAck As Boolean     ' own traffic needing the addressees' acknowledgement
Private m_firstLoad As Boolean  ' Form_Timer's first tick loads the TempVar's message

Public Property Get HelpTopic() As String
    HelpTopic = "message"
End Property

Private Sub Form_Load()
    Dim c As Control
    On Error GoTo Err_Handler
    m_content = ";"
    For Each c In Me.Controls
        Select Case True
            Case c.Name = "txtFocusSink", c.Name = "barTitle", c.Name = "lblTitle", c.Name = "lblSubtitle", c.Name = "tipTitle", _
                    c.Name = "btnBack", c.Name Like "emptyMsg*"
            Case c.Visible
                m_content = m_content & c.Name & ";"
        End Select
    Next
    m_tab = "d"
    SelectTab Me, m_tab
    For Each c In Me.Controls               ' nothing shows until Form_Timer has loaded the message
        If InStr(m_content, ";" & c.Name & ";") > 0 Then SetVisible c, False
    Next
    m_firstLoad = True
    Me.TimerInterval = 1
    Exit Sub
Err_Handler:
    HandleError "pgMessage.Form_Load"
End Sub

' Once loaded, the TempVar's message; later, a related message picked, now
' the click has finished.
Private Sub Form_Timer()
    Dim nextId As Long
    On Error GoTo Err_Handler
    If Me.TimerInterval = 0 Then Exit Sub   ' a tick left over from before
    Me.TimerInterval = 0
    If m_firstLoad Then
        m_firstLoad = False
        m_nextId = 0
        LoadMessage Nz(TempVars("OpenMessageID"), 0)
        Exit Sub
    End If
    nextId = m_nextId
    m_nextId = 0
    If nextId <> 0 Then LoadMessage nextId
    Exit Sub
Err_Handler:
    HandleError "pgMessage.Form_Timer", m_nextId
End Sub

' Load message nextId from the timer (see Form_Timer).
Private Sub OpenLater(ByVal nextId As Long)
    If nextId = 0 Or nextId = m_id Or m_firstLoad Then Exit Sub
    TempVars("OpenMessageID") = nextId
    m_nextId = nextId
    Me.TimerInterval = 10
End Sub

' ---------------------------------------------------------------------------
' Tabs and the text switch (the kit's builders call these)
' ---------------------------------------------------------------------------

Public Function TabPick(ByVal key As String)
    On Error GoTo Err_Handler
    If key = m_tab Then Exit Function
    m_tab = key
    SelectTab Me, key
    ShowTab
    Exit Function
Err_Handler:
    HandleError "pgMessage.TabPick", key
End Function

' As formatted / Raw as received (segText's onChange).
Public Sub TextViewChanged()
    On Error GoTo Err_Handler
    ShowTextView
    Exit Sub
Err_Handler:
    HandleError "pgMessage.TextViewChanged"
End Sub

' ---------------------------------------------------------------------------
' Buttons, links and rows
' ---------------------------------------------------------------------------

Private Sub btnBack_Click()
    On Error GoTo Err_Handler
    NavTo "messages"
    Exit Sub
Err_Handler:
    HandleError "pgMessage.btnBack_Click"
End Sub

' The empty state's button.
Private Sub emptyMsg_b_Click()
    On Error GoTo Err_Handler
    NavTo "messages"
    Exit Sub
Err_Handler:
    HandleError "pgMessage.emptyMsg_b_Click"
End Sub

' Record the person's acknowledgement (designated traffic only). The right
' and the need are checked again in the data file (RequireRole, and
' Acknowledge raises when it is no longer needed).
Private Sub btnAck_Click()
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    RequireRole "VIEW", "acknowledge messages"
    Acknowledge m_id
    KitToast "Acknowledged " & MessageRef(m_id) & ".", "success"
    ShowAck
    Exit Sub
Err_Handler:
    HandleError "pgMessage.btnAck_Click", m_id
End Sub

Private Sub btnUnread_Click()
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    MarkUnread m_id
    KitToast "Marked unread. It shows with a dot under Messages."
    Exit Sub
Err_Handler:
    HandleError "pgMessage.btnUnread_Click", m_id
End Sub

' The official rendition, to print: audited, then previewed.
Private Sub btnPrint_Click()
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    AuditTry "Print", "Message (" & MessageRef(m_id) & ")", "Official rendition opened to print, with its integrity block.", _
        MessageClassification(m_id)
    DoCmd.OpenReport RPT_MESSAGE, acViewPreview, , , , "id=" & m_id
    Exit Sub
Err_Handler:
    If Err.Number = ERR_REPORT_CANCELLED Then Exit Sub   ' the report said why it didn't open
    HandleError "pgMessage.btnPrint_Click", m_id
End Sub

' The lifecycle banner's link: the replacing or related message.
Private Sub lnkLife_Click()
    On Error GoTo Err_Handler
    If Len(m_lifeRef) > 0 Then OpenLater MessageIdByRef(m_lifeRef)
    Exit Sub
Err_Handler:
    HandleError "pgMessage.lnkLife_Click", m_lifeRef
End Sub

' A row clicked: a related message opens here.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    If grid = "sfRel" Then OpenLater Nz(Me.sfRel.Form.txtSel, 0)
    Exit Sub
Err_Handler:
    HandleError "pgMessage.GridClick", grid
End Sub

' ---------------------------------------------------------------------------
' The message
' ---------------------------------------------------------------------------

' Show message msgId, marking it read; or say why there is none.
Private Sub LoadMessage(ByVal msgId As Long)
    Dim rs As DAO.Recordset, s As String, ref As String, vrnt As String, life As String, sql As String
    m_id = 0
    If Not DataLinked() Then
        ShowMissing "No data file linked", "This copy was built without the data file, so there is no message to show."
        Exit Sub
    End If
    If msgId = 0 Then
        ShowMissing "No message picked", "Open a message from Home, Messages or Search to read it here."
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblMessages WHERE MessageID=" & msgId, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ShowMissing "No message to show", "ANCHOR does not hold this message. It may not have been ingested yet."
        Exit Sub
    End If
    m_id = msgId
    ShowContent True
    MarkRead msgId

    ' The header.
    Me.lblSubtitle.Caption = NoAmp(Nz(rs!Reference, "") & "  |  " & Nz(rs!Originator, "") & "  |  " & Nz(rs!Dtg, "") & _
        IIf(IsNull(rs!MsgType), "", "  |  " & rs!MsgType))
    Me.lblSubject.Caption = NoAmp(Nz(rs!Subject, ""))
    SetPill Me.pilClass, rs!Classification
    SetPill Me.pilPrec, rs!Precedence
    SetPill Me.pilLife, rs!Lifecycle
    Me.lblMarking.Caption = IIf(IsMarkedAboveUnclass(rs!Classification), NoAmp(ClassificationMarking(rs!Classification)), "")

    ' The lifecycle banner and its link.
    life = UCase$(Nz(rs!Lifecycle, ""))
    s = LifecycleNotice(life, rs!ReplacedBy, rs!RelatesTo)
    Select Case life
        Case "SUPERSEDED", "CANCELLED", "AMENDED": ref = Nz(rs!ReplacedBy, "")
        Case "CORRECTED": ref = Nz(rs!RelatesTo, "")
        Case "ACTIVE"
            ref = Nz(rs!RelatesTo, "")
            If Len(ref) > 0 Then s = "Current. This message supersedes or updates " & ref & "." Else s = "Current. This is the version in force."
    End Select
    Select Case life
        Case "ACTIVE": vrnt = IIf(Len(ref) > 0, "info", "success")
        Case "CANCELLED": vrnt = "destructive"
        Case "SUPERSEDED", "EXPIRED": vrnt = "warning"
        Case Else: vrnt = "info"
    End Select
    If Len(s) = 0 Then s = "Status: " & life & "."
    ShowAlert Me, "alLife", NoAmp(s), vrnt
    m_lifeRef = ""
    If Len(ref) > 0 Then
        If MessageIdByRef(ref) <> 0 Then m_lifeRef = ref
    End If
    Me.lnkLife.Caption = NoAmp("Open " & m_lifeRef)
    SetVisible Me.lnkLife, Len(m_lifeRef) > 0

    ' The official text, the raw text and the AI summary.
    If Len(Nz(rs!BodyText, "")) > 0 Then
        Me.txtBody = ShowLines(rs!BodyText)
    Else
        Me.txtBody = "(ANCHOR holds no formatted text for this message: switch to Raw as received.)"
    End If
    Me.txtRaw = ShowLines(rs!RawText)
    ShowTextView
    If Len(Nz(rs!AiSummary, "")) > 0 Then
        Me.txtAi = rs!AiSummary
    Else
        Me.txtAi = "No AI summary is stored for this message."
    End If

    ' Details.
    s = Nz(DLookup("TypeName", "tblMessageTypes", "TypeCode=" & Q(rs!MsgType)), "")
    Me.txtType = Nz(rs!MsgType, "Not recorded") & IIf(Len(s) > 0, " (" & s & ")", "")
    s = MsgClassText(rs!MsgClass)
    Me.txtClass = IIf(Len(s) > 0, s, "Not recorded")
    Me.txtReceived = Format$(rs!ReceivedOn, "dd mmm yyyy hh:nn")
    Me.txtC2oix = Nz(rs!C2oixId, "Not recorded")
    Me.txtSha = Nz(rs!Sha256, "Not recorded")
    Me.txtSize = IIf(IsNull(rs!SizeKb), "Not recorded", Format$(rs!SizeKb, "#,##0") & " KB (" & Format$(Nz(rs!SizeKb, 0) / 1024, "0.0") & " MB)")
    Me.txtRetention = Nz(rs!Retention, "Not recorded")
    If Len(Nz(rs!ActionText, "")) > 0 Then
        Me.txtAction = rs!ActionText & vbCrLf & "Suspense: " & Nz(rs!ActionSuspense, "none") & vbCrLf & _
            "Assigned to: " & Nz(rs!ActionAssignedTo, "no one named")
    Else
        Me.txtAction = "No action item."
    End If

    ' Addressees, distribution and attachments.
    Me.sfAddr.Form.RecordSource = "SELECT AddresseeID AS id, AddrName AS addr_name, AddrRole AS addr_role, AddrKind AS addr_kind " & _
        "FROM tblAddressees WHERE MessageID=" & msgId & " ORDER BY IIf(AddrRole='TO',0,1), AddrName"
    Me.txtDist = IIf(Len(Nz(rs!Distribution, "")) > 0, rs!Distribution, "None recorded.")
    Me.sfAtt.Form.RecordSource = "SELECT AttachmentID AS id, FileName AS att_name, SizeText AS att_size FROM tblAttachments " & _
        "WHERE MessageID=" & msgId & " ORDER BY FileName"
    m_attN = Me.sfAtt.Form.RecordsetClone.RecordCount

    ' Related messages.
    sql = RelatedMessagesSql(msgId)
    If Len(sql) > 0 Then
        Me.sfRel.Form.RecordSource = sql
        m_relN = Me.sfRel.Form.RecordsetClone.RecordCount
    Else
        m_relN = 0
    End If
    Me.sfRel.Form.txtSel = Null
    rs.Close

    ShowAck
    ShowTab
End Sub

' The acknowledgement bar, and the Acks tab's summary and list.
Private Sub ShowAck()
    Dim total As Long, acked As Long, ackOn As Variant, req As Boolean
    If m_id = 0 Then Exit Sub
    req = Nz(DLookup("AckRequired", "tblMessages", "MessageID=" & m_id), False)
    m_ownAck = req And SentByOwnCommand(m_id)
    If m_ownAck Then
        AckStatusCounts m_id, total, acked
        ShowAlert Me, "alAck", "Sent by your command: " & acked & " of " & total & " addressee " & Plural(total, "command") & _
            " acknowledged (Acks tab).", "info"
        SetVisible Me.btnAck, False
        Me.lblAckSum.Caption = acked & " of " & total & " addressee " & Plural(total, "command") & " acknowledged. " & _
            "Groups count as their commands."
        Me.sfAcks.Form.RecordSource = AckStatusSql(m_id)
    ElseIf NeedsOwnAck(m_id) Then
        ShowAlert Me, "alAck", NoAmp("Your acknowledgement is required. " & _
            AckDueText(DLookup("AckDue", "tblMessages", "MessageID=" & m_id))), "destructive"
        SetVisible Me.btnAck, True
        Me.lblAckSum.Caption = NoAmp("Only traffic " & CommandSource() & " sends is tracked here. Yours is in the bar above.")
    Else
        ackOn = AckOnOf(m_id)
        If IsNull(ackOn) Then
            ShowAlert Me, "alAck", "No acknowledgement is required of you.", "info"
        Else
            ShowAlert Me, "alAck", "Acknowledged " & Format$(ackOn, "dd mmm yyyy hh:nn") & ".", "success"
        End If
        SetVisible Me.btnAck, False
        Me.lblAckSum.Caption = NoAmp("Only traffic " & CommandSource() & " sends is tracked here.")
    End If
    If m_tab = "k" Then ShowTab
End Sub

' Show the open tab's controls and hide the others', then what the open
' tab has no rows for.
Private Sub ShowTab()
    Dim g As Variant, nm As Variant, isOpen As Boolean
    If m_id = 0 Then Exit Sub
    For Each g In Array("d", "a", "k", "r")
        isOpen = (g = m_tab)
        For Each nm In Split(Me.Controls("grp" & UCase$(g)).Caption, ";")
            If Len(nm) > 0 Then SetVisible Me.Controls(nm), isOpen
        Next
    Next
    If m_tab = "a" Then
        SetVisible Me.sfAtt, m_attN > 0
        For Each nm In Split(Me.grpA.Caption, ";")
            If nm Like "hdrAtt*" Then SetVisible Me.Controls(nm), m_attN > 0
        Next
    End If
    SetVisible Me.lblAttNone, m_tab = "a" And m_attN = 0
    If m_tab = "k" And Not m_ownAck Then
        For Each nm In Split(Me.grpK.Caption, ";")
            If nm <> "lblAckSum" Then SetVisible Me.Controls(nm), False
        Next
    End If
    If m_tab = "r" And m_relN = 0 Then
        For Each nm In Split(Me.grpR.Caption, ";")
            SetVisible Me.Controls(nm), False
        Next
    End If
    SetVisible Me.lblRelNone, m_tab = "r" And m_relN = 0
End Sub

' The official text or the raw text, as the switch says.
Private Sub ShowTextView()
    Dim raw As Boolean
    raw = (SegmentValue(Me, "segText") = "raw")
    SetVisible Me.txtBody, Not raw
    SetVisible Me.txtRaw, raw
End Sub

' Hide everything but the header and say why there is no message.
Private Sub ShowMissing(ByVal title As String, ByVal text As String)
    ShowContent False
    Me.lblSubtitle.Caption = ""
    SetVisible Me.txtRaw, False
    SetVisible Me.lblAttNone, False
    SetVisible Me.lblRelNone, False
    ShowEmpty Me, "emptyMsg", True, title, text
End Sub

' Show or hide every control listed in m_content (the empty state the other way).
Private Sub ShowContent(ByVal isShown As Boolean)
    Dim nm As Variant
    For Each nm In Split(m_content, ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), isShown
    Next
    ShowEmpty Me, "emptyMsg", Not isShown
End Sub

' Put a word in a pill (PillBox) and repaint its rounded image.
Private Sub SetPill(ByVal c As Control, ByVal v As Variant)
    c.Value = Nz(v, "")
    Me.Controls(c.Name & "_bg").Requery
End Sub

' Text for a label's caption, where & would underline the next letter.
Private Function NoAmp(ByVal s As String) As String
    NoAmp = Replace(s, "&", "&&")
End Function
