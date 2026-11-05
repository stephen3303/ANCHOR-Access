Option Compare Database
Option Explicit

' Briefings (see modBuildBriefings.Build_pgBriefings). The saved briefings
' (tblBriefings) on the left; picking one opens it in the builder, New
' briefing starts an empty one.
'
' The builder: a title and the period (From and To, whole days). "Add
' matching messages" adds the period's messages that match the precedence,
' message type (tblMessageTypes) and community filters to the pick list,
' not ticked (modBriefings.BrfAddMatching); adding again with other filters
' adds more. Tick / untick, or a double-click, picks a message to brief;
' Note... writes what the briefing says about it (a message with a note is
' ticked). The notes box is for the whole briefing. Save (Ctrl+S) writes the
' briefing and its ticked messages and audits it (BrfSave checks the
' Briefing Creator role in the data file). Print briefing opens rptBriefing
' for the saved briefing, saving first when there are changes.
'
' Changes not saved yet are asked about before leaving the page or opening
' another briefing (CanLeave).
'
' Depends on: modBriefings, modGlobals, modMessages, modSignIn, modUI; the
' kit's modKitKeys, modKitUI.

Private Const MSG_TITLE As String = "Briefings"

Private m_briefId As Long       ' the briefing open in the builder (0: a new one)
Private m_isDirty As Boolean    ' it has changes not saved yet
Private m_hasData As Boolean    ' a data file is linked

Public Property Get HelpTopic() As String
    HelpTopic = "briefings"
End Property

Private Sub Form_Load()
    Dim c As Variant
    On Error GoTo Err_Handler
    m_hasData = DataLinked()
    If Not m_hasData Then
        For Each c In Array("btnNew", "btnAdd", "btnTick", "btnNote", "btnSave", "btnPrint", "txtTitle", "txtComm", "txtNotes", _
                "cboPrec", "cboType", "dtFrom", "dtTo", "cal_dtFrom", "cal_dtTo")
            SetEnabled Me.Controls(CStr(c)), False
        Next
        ShowEmpty Me, "emptySaved", True, "No data file linked", "There are no briefings to show."
        SetVisible Me.sfSaved, False
        ShowEmpty Me, "emptyPick", True, "No data file linked", "This copy was built without the data file, so there is no traffic to brief."
        SetVisible Me.sfPick, False
        Exit Sub
    End If
    Me.cboPrec.RowSource = BrfPrecedenceList()
    Me.cboType.RowSource = BrfTypeList()
    Me.sfSaved.Form.RecordSource = BrfSavedSql()
    ShowSavedEmpty
    NewBriefing
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.Form_Load"
End Sub

' The shell asks before leaving the page (ShellCanLeave).
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = AskToSave()
    Exit Function
Err_Handler:
    HandleError "pgBriefings.CanLeave"
End Function

' Ctrl+S saves the briefing, Ctrl+N starts a new one (modKitKeys).
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    If Not m_hasData Then Exit Function
    Select Case action
        Case SHORTCUT_SAVE
            SaveBriefing
            Shortcut = True
        Case SHORTCUT_NEW
            If AskToSave() Then NewBriefing
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgBriefings.Shortcut", action
    Shortcut = True
End Function

' True when the open briefing may be put away: no changes, or the user saved
' them or chose to drop them. False when they cancelled or the save failed.
Private Function AskToSave() As Boolean
    If Not m_isDirty Then AskToSave = True: Exit Function
    Select Case KitMsgBox("Save your changes to this briefing?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: AskToSave = SaveBriefing()
        Case vbNo: AskToSave = True
        Case Else: AskToSave = False
    End Select
End Function

' ---------------------------------------------------------------------------
' Opening a briefing
' ---------------------------------------------------------------------------

' An empty briefing: no title, the last 7 days, every precedence and type.
Private Sub NewBriefing()
    m_briefId = 0
    Me.txtTitle = Null
    Me.dtFrom = Date - 7
    Me.dtTo = Date
    Me.cboPrec = "*"
    Me.cboType = "*"
    Me.txtComm = Null
    Me.txtNotes = Null
    Me.lblBuildHdr.Caption = "New briefing"
    BrfClearPick
    Me.sfSaved.Form.txtSel = Null
    Me.sfSaved.Form.Recalc
    m_isDirty = False
    RefreshPick
End Sub

' Open saved briefing briefId in the builder.
Private Sub LoadBriefing(ByVal briefId As Long)
    Dim rs As DAO.Recordset, crit As String
    Set rs = CurrentDb.OpenRecordset("SELECT Title, PeriodFrom, PeriodTo, Criteria, Notes FROM tblBriefings WHERE BriefingID=" & briefId, _
        dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        KitMsgBox "That briefing was deleted. The list is up to date now.", vbInformation, MSG_TITLE
        Me.sfSaved.Form.Requery
        ShowSavedEmpty
        NewBriefing
        Exit Sub
    End If
    m_briefId = briefId
    Me.txtTitle = rs!Title
    Me.dtFrom = rs!PeriodFrom
    Me.dtTo = rs!PeriodTo
    crit = Nz(rs!Criteria, "")
    Me.txtNotes = rs!Notes
    rs.Close
    Me.cboPrec = BrfCriteriaPart(crit, "prec", "*")
    Me.cboType = BrfCriteriaPart(crit, "type", "*")
    If IsNull(Me.cboType) Then Me.cboType = "*"     ' a type no longer on the list
    Me.txtComm = BrfCriteriaPart(crit, "community", "")
    Me.lblBuildHdr.Caption = Replace("Briefing: " & ClipText(Nz(Me.txtTitle, ""), 60), "&", "&&")
    BrfLoadPick briefId
    m_isDirty = False
    RefreshPick
End Sub

Private Sub ShowSavedEmpty()
    Dim n As Long
    n = Me.sfSaved.Form.RecordsetClone.RecordCount
    ShowEmpty Me, "emptySaved", n = 0
    SetVisible Me.sfSaved, n > 0
End Sub

' ---------------------------------------------------------------------------
' The pick list
' ---------------------------------------------------------------------------

' Reload the pick list, keeping the picked row, and its counts.
Private Sub RefreshPick()
    Dim keep As Variant, nListed As Long, nTicked As Long
    keep = Null
    If Len(Me.sfPick.Form.RecordSource) > 0 Then keep = Me.sfPick.Form.txtSel
    Me.sfPick.Form.RecordSource = BrfPickSql()
    BrfPickCounts nListed, nTicked
    Me.lblCount.Caption = nListed & " listed, " & nTicked & " ticked to brief"
    ShowEmpty Me, "emptyPick", nListed = 0
    SetVisible Me.sfPick, nListed > 0
    SetEnabled Me.btnTick, nListed > 0
    SetEnabled Me.btnNote, nListed > 0
    If Not GridSelect(Me.sfPick, keep) Then SelectFirstRow Me.sfPick
End Sub

' The message picked on the pick list, 0 for none.
Private Function PickedId() As Long
    PickedId = Nz(Me.sfPick.Form.txtSel, 0)
End Function

' The period's dates, or False after saying what is wrong with them.
Private Function PeriodOk() As Boolean
    If Not IsDate(Me.dtFrom) Or Not IsDate(Me.dtTo) Then
        KitMsgBox "Enter both dates of the period.", vbExclamation, MSG_TITLE
    ElseIf DateValue(Me.dtFrom) > DateValue(Me.dtTo) Then
        KitMsgBox "The From date is after the To date. Swap them.", vbExclamation, MSG_TITLE
    Else
        PeriodOk = True
    End If
End Function

Private Sub btnAdd_Click()
    Dim n As Long
    On Error GoTo Err_Handler
    If Not PeriodOk() Then Exit Sub
    n = BrfAddMatching(DateValue(Me.dtFrom), DateValue(Me.dtTo), Nz(Me.cboPrec, "*"), Nz(Me.cboType, "*"), _
        FieldText(Me, Me.txtComm))
    RefreshPick
    If n = 0 Then
        KitToast "No more messages match. Widen the period or the filters.", "info"
    Else
        KitToast "Added " & n & " " & Plural(n, "message") & ". Tick the ones to brief.", "success"
    End If
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.btnAdd_Click"
End Sub

Private Sub btnTick_Click()
    On Error GoTo Err_Handler
    TickPicked
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.btnTick_Click"
End Sub

' Brief the picked message, or stop briefing it.
Private Sub TickPicked()
    If PickedId() = 0 Then Exit Sub
    BrfToggle PickedId()
    m_isDirty = True
    RefreshPick
End Sub

Private Sub btnNote_Click()
    Dim v As Variant, msgId As Long
    On Error GoTo Err_Handler
    msgId = PickedId()
    If msgId = 0 Then Exit Sub
    v = KitInputBox("What should the briefing say about " & MessageRef(msgId) & "? Leave it blank to brief the first " & _
        "paragraph of its official text.", "Note", BrfNoteOf(msgId))
    If IsNull(v) Then Exit Sub
    BrfSetNote msgId, CStr(v)
    m_isDirty = True
    RefreshPick
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.btnNote_Click"
End Sub

' A row clicked: a saved briefing opens; a double-click on the pick list
' ticks or unticks the message.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    Select Case grid
        Case "sfSaved"
            id = Nz(Me.sfSaved.Form.txtSel, 0)
            If id = 0 Or id = m_briefId Then Exit Sub
            If AskToSave() Then
                LoadBriefing id
            Else
                GridSelect Me.sfSaved, m_briefId     ' stay on the open one
            End If
        Case "sfPick"
            If dbl Then TickPicked
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.GridClick", grid
End Sub

' ---------------------------------------------------------------------------
' Changes, Save and Print
' ---------------------------------------------------------------------------

Private Sub txtTitle_AfterUpdate()
    On Error GoTo Err_Handler
    m_isDirty = True
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.txtTitle_AfterUpdate"
End Sub

Private Sub txtNotes_AfterUpdate()
    On Error GoTo Err_Handler
    m_isDirty = True
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.txtNotes_AfterUpdate"
End Sub

Private Sub dtFrom_AfterUpdate()
    On Error GoTo Err_Handler
    m_isDirty = True
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.dtFrom_AfterUpdate"
End Sub

Private Sub dtTo_AfterUpdate()
    On Error GoTo Err_Handler
    m_isDirty = True
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.dtTo_AfterUpdate"
End Sub

' The calendar buttons: the kit's picker sets the box without its
' AfterUpdate, so the change is marked here.
Private Sub cal_dtFrom_Click()
    On Error GoTo Err_Handler
    PickDate Me, "dtFrom"
    m_isDirty = True
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.cal_dtFrom_Click"
End Sub

Private Sub cal_dtTo_Click()
    On Error GoTo Err_Handler
    PickDate Me, "dtTo"
    m_isDirty = True
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.cal_dtTo_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    If AskToSave() Then NewBriefing
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.btnNew_Click"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveBriefing
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.btnSave_Click"
End Sub

' Save the open briefing (BrfSave). True when it saved; False after saying
' what is missing, or when the save failed (the error shown).
Private Function SaveBriefing() As Boolean
    Dim title As String, nListed As Long, nTicked As Long
    On Error GoTo Fail
    title = Trim$(FieldText(Me, Me.txtTitle))
    If Len(title) = 0 Then
        KitMsgBox "Give the briefing a title.", vbExclamation, MSG_TITLE
        Exit Function
    End If
    If Not PeriodOk() Then Exit Function
    BrfPickCounts nListed, nTicked
    If nTicked = 0 Then
        KitMsgBox "Tick at least one message to brief.", vbExclamation, MSG_TITLE
        Exit Function
    End If
    m_briefId = BrfSave(m_briefId, title, DateValue(Me.dtFrom), DateValue(Me.dtTo), _
        BrfCriteriaText(Nz(Me.cboPrec, "*"), Nz(Me.cboType, "*"), FieldText(Me, Me.txtComm)), FieldText(Me, Me.txtNotes))
    m_isDirty = False
    Me.lblBuildHdr.Caption = Replace("Briefing: " & ClipText(title, 60), "&", "&&")
    GridSelect Me.sfSaved, m_briefId, True
    ShowSavedEmpty
    KitToast "Briefing saved with " & nTicked & " " & Plural(nTicked, "message") & ".", "success"
    SaveBriefing = True
    Exit Function
Fail:
    HandleError "pgBriefings.SaveBriefing", title
End Function

' Print the saved briefing; one with changes, or never saved, is saved first.
Private Sub btnPrint_Click()
    On Error GoTo Err_Handler
    If m_isDirty Or m_briefId = 0 Then
        If KitMsgBox("The briefing is printed as saved. Save it now and print it?", vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub
        If Not SaveBriefing() Then Exit Sub
    End If
    BrfPrint m_briefId
    Exit Sub
Err_Handler:
    HandleError "pgBriefings.btnPrint_Click"
End Sub
