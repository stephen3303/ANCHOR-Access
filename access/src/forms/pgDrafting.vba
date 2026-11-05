Option Compare Database
Option Explicit

' Drafting & Release (see modBuildDrafting.Build_pgDrafting): the drafts,
' the latest change first (modDrafting.DraftListSql), filtered by a chip,
' the search box and the column funnels, and the picked one on the right in
' four tabs:
'   Draft     the header and the text; Save and Undo while its drafter may
'             edit it (drafting or returned), Withdraw before release
'   Message   the message as it transmits (placeholders for the DTG and
'             sequence until release, the frozen text after) and the naval
'             format check, run on the text as it would go if released now
'   Route     where it stands, the chain it took, and the action the
'             signed-in person may take: Submit (its drafter), Endorse or
'             Release and Return (the holder of the current step)
'   Log       the routing log
' Every action is checked again in the data file when it is taken
' (modDrafting); a refusal shows its sentence. New draft opens dlgNewDraft
' (blank, or from an active template, whose fields are asked for in turn).
' Another page shows a draft first through TempVars("DraftingDraftID").
'
' Owner: drafting and release. Depends on: modDrafting, modErrors,
' modMessages, modNavalFormat, modRouting, modSignIn, modTemplates, modUI;
' the kit's modKitUI, modKitUtil, modGridFilter.

Private Const MSG_TITLE As String = "Drafting"

Private m_chip As Long          ' DRAFT_CHIPS; 0 All
Private m_search As String
Private m_id As Long            ' the draft shown (0: none)
Private m_tab As String         ' draft, message, route or log
Private m_d As Object           ' its values as last read (DraftValues)
Private m_sig As String         ' the editor's values as loaded (DirtyNow compares)
Private m_editable As Boolean   ' the signed-in person may edit it

Public Property Get HelpTopic() As String
    HelpTopic = "drafting"
End Property

Private Sub Form_Load()
    Dim openId As Long
    On Error GoTo Err_Handler
    m_tab = "draft"
    ChipOn Me, "chip", 0
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no drafts."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        SetEnabled Me.btnNew, False
        ShowDraft 0
        Exit Sub
    End If
    Me.cboType.RowSource = "SELECT TypeCode, TypeName FROM tblMessageTypes WHERE IsActive ORDER BY SortOrder"
    Me.cboClass.RowSource = "SELECT ClassCode, ClassName FROM tblMessageClasses ORDER BY ClassCode"
    openId = Nz(TempVars("DraftingDraftID"), 0)
    TempVars("DraftingDraftID") = Null
    RefreshList openId
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.Form_Load"
End Sub

' The shell asks before it leaves the page: unsaved edits are saved,
' dropped or keep the person here.
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = OfferSave()
    Exit Function
Err_Handler:
    HandleError "pgDrafting.CanLeave"
    CanLeave = True
End Function

' Ctrl+S saves, Ctrl+N starts a draft, Ctrl+F finds.
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            If m_editable And m_id <> 0 Then SaveClicked
            Shortcut = True
        Case SHORTCUT_NEW
            If Me.btnNew.Enabled Then btnNew_Click
            Shortcut = True
        Case SHORTCUT_FIND
            Me.txtSearch.SetFocus
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgDrafting.Shortcut", action
    Shortcut = True
End Function

' ---------------------------------------------------------------------------
' Chips, search and tabs (the kit's builders call these)
' ---------------------------------------------------------------------------

Public Function ChipPick(ByVal i As Long)
    On Error GoTo Err_Handler
    If i = m_chip Then Exit Function
    If Not OfferSave() Then Exit Function
    m_chip = i
    ChipOn Me, "chip", i
    RefreshList m_id
    Exit Function
Err_Handler:
    HandleError "pgDrafting.ChipPick", i
End Function

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList m_id
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.SearchFor", text
End Sub

Public Sub ListFiltered(ByVal sfName As String)
    On Error GoTo Err_Handler
    AfterFilter
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.ListFiltered", sfName
End Sub

Public Function TabPick(ByVal key As String)
    On Error GoTo Err_Handler
    m_tab = key
    ShowTab
    Exit Function
Err_Handler:
    HandleError "pgDrafting.TabPick", key
End Function

' ---------------------------------------------------------------------------
' Rows
' ---------------------------------------------------------------------------

Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    If grid <> "sfList" Then Exit Sub
    id = Nz(Me.sfList.Form.txtSel, 0)
    If id = m_id Then Exit Sub
    If Not OfferSave() Then GridSelect Me.sfList, m_id: Exit Sub
    ShowDraft id
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.GridClick", grid & " " & id
End Sub

' The empty state's button: the All chip, no search, no column filters.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    m_chip = 0
    ChipOn Me, "chip", 0
    ClearGridFilters Me.sfList
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.emptyList_b_Click"
End Sub

' ---------------------------------------------------------------------------
' Buttons
' ---------------------------------------------------------------------------

Private Sub btnNew_Click()
    Dim tid As Long, id As Long, problem As String
    On Error GoTo Err_Handler
    If Not OfferSave() Then Exit Sub
    TempVars("NewDraftTemplateID") = Null
    If Not RunDialog("dlgNewDraft") Then Exit Sub
    tid = Nz(TempVars("NewDraftTemplateID"), 0)
    TempVars("NewDraftTemplateID") = Null
    If tid = 0 Then
        id = NewBlankDraft()
    Else
        id = StartDraftFromTemplate(tid, problem)
        If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: Exit Sub
        If id = 0 Then Exit Sub                 ' cancelled while filling in the fields
    End If
    If m_chip = 1 Or m_chip = 3 Then m_chip = 0: ChipOn Me, "chip", 0
    m_tab = "draft"
    RefreshList id
    KitToast "Draft started. Save it as you go, then submit it on the Route tab.", "success"
    Me.txtSubject.SetFocus
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnNew_Click", tid
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveClicked
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnSave_Click", m_id
End Sub

Private Sub btnUndo_Click()
    On Error GoTo Err_Handler
    ShowDraft m_id
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnUndo_Click", m_id
End Sub

Private Sub cboClassif_AfterUpdate()
    On Error GoTo Err_Handler
    ShowMarking
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.cboClassif_AfterUpdate"
End Sub

Private Sub btnWithdraw_Click()
    Dim why As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    If KitMsgBox("Withdraw this draft? It stops here and is kept for the record; nobody can submit or release it again.", _
            vbYesNo + vbExclamation, MSG_TITLE) <> vbYes Then Exit Sub
    why = WithdrawDraft(m_id)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE Else KitToast "Draft withdrawn.", "success"
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnWithdraw_Click", m_id
End Sub

Private Sub btnCheck_Click()
    On Error GoTo Err_Handler
    RunCheck
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnCheck_Click", m_id
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If KitCopyText(Nz(Me.txtPreview, "")) Then KitToast "Message text copied.", "success"
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnCopy_Click", m_id
End Sub

Private Sub btnSubmit_Click()
    Dim why As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    If DirtyNow() Then
        If Not SaveClicked() Then Exit Sub
    End If
    why = SubmitDraft(m_id)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: ShowDraft m_id: Exit Sub
    KitToast "Submitted. It waits on " & CurrentHolderText() & ".", "success"
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnSubmit_Click", m_id
End Sub

Private Sub btnEndorse_Click()
    Dim why As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    why = EndorseDraft(m_id, Nz(Me.txtComment, ""))
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: RefreshList m_id: Exit Sub
    KitToast "Endorsed. It now waits on " & CurrentHolderText() & ".", "success"
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnEndorse_Click", m_id
End Sub

Private Sub btnReturn_Click()
    Dim why As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    If Len(Trim$(Nz(Me.txtComment, ""))) = 0 Then
        KitMsgBox "Type a comment saying what to correct, then Return.", vbExclamation, MSG_TITLE
        Me.txtComment.SetFocus
        Exit Sub
    End If
    why = ReturnDraft(m_id, Nz(Me.txtComment, ""))
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: RefreshList m_id: Exit Sub
    KitToast "Returned to " & m_d("drafter_name") & " for correction.", "success"
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnReturn_Click", m_id
End Sub

Private Sub btnRelease_Click()
    Dim why As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    why = ReleaseProblem(m_id)
    If Len(why) = 0 Then why = FormatIssuesOf(DraftFormatIssues(m_d), False)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: Exit Sub
    If KitMsgBox("Release this message now? It takes today's DTG and sequence " & LocalSequenceText(SerialText(NextSequenceNumber())) & _
            ", its text is frozen, and it joins Messages. This can't be undone.", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
    why = ReleaseDraft(m_id, Nz(Me.txtComment, ""))
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: RefreshList m_id: Exit Sub
    Set m_d = DraftValues(m_id)
    KitToast "Released " & m_d("dtg") & " as " & m_d("local_seq") & ". It is in Messages now.", "success"
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnRelease_Click", m_id
End Sub

Private Sub btnOpenMessage_Click()
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    If m_d("message_id") <> 0 Then OpenMessage m_d("message_id")
    Exit Sub
Err_Handler:
    HandleError "pgDrafting.btnOpenMessage_Click", m_id
End Sub

' ---------------------------------------------------------------------------
' The list
' ---------------------------------------------------------------------------

' Load the drafts for the chip and the search, with each chip's count, and
' show openId (0, or not listed: the first).
Private Sub RefreshList(ByVal openId As Long)
    Dim i As Long
    If Not DataLinked() Then Exit Sub
    Me.sfList.Form.RecordSource = DraftListSql(m_chip, m_search)
    SetGridFilter Me.sfList, ""
    GridHighlight Me.sfList, m_search
    For i = 0 To UBound(Split(DRAFT_CHIPS, "|"))
        StyleChip Me.Controls("chip_" & i), i = m_chip, DraftCount(i)
    Next
    m_id = openId
    m_sig = ""                                  ' what was shown is reloaded, not kept
    AfterFilter
End Sub

Private Sub AfterFilter()
    Dim n As Long, keep As Long
    n = Me.sfList.Form.RecordsetClone.RecordCount
    If DraftCount(0) = 0 Then
        ShowEmpty Me, "emptyList", n = 0, "No drafts yet", "New draft starts one, blank or from a template."
        SetVisible Me.emptyList_b, False
    Else
        ShowEmpty Me, "emptyList", n = 0, "No drafts match", "No draft fits the chip, the search and the column filters."
        Me.emptyList_b.Caption = "Show all"
    End If
    SetVisible Me.sfList, n > 0
    keep = m_id
    If GridSelect(Me.sfList, keep) Then
        ShowDraft keep
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ShowDraft 0
    Else
        ShowDraft Nz(Me.sfList.Form.txtSel, 0)
    End If
End Sub

' ---------------------------------------------------------------------------
' The draft
' ---------------------------------------------------------------------------

' Show draft draftId on the right (0: the hint instead), on the open tab.
Private Sub ShowDraft(ByVal draftId As Long)
    Dim c As Control
    m_id = draftId
    Set m_d = Nothing
    If draftId <> 0 Then
        Set m_d = DraftValues(draftId)
        If m_d Is Nothing Then m_id = 0
    End If
    If m_id <> 0 Then
        m_editable = (Len(DraftEditProblem(m_id)) = 0)
        Me.lblDraftRef.Caption = DraftStatusWord(m_d("status")) & " " & ChrW(183) & " " & m_d("msg_type") & _
            IIf(Len(m_d("local_seq")) > 0, " " & ChrW(183) & " " & m_d("local_seq") & " " & m_d("dtg"), "") & _
            " " & ChrW(183) & " drafted by " & m_d("drafter_name")
        Me.txtSubject = m_d("subject")
        Me.cboType = m_d("msg_type")
        Me.cboClass = m_d("msg_class")
        Me.cboPrec = m_d("precedence")
        Me.cboClassif = m_d("classification")
        Me.txtOriginator = m_d("originator")
        Me.txtTo = m_d("to_line")
        Me.txtInfo = m_d("info_line")
        Me.txtParas = m_d("paragraphs")
        Me.txtComment = Null
        Me.sfSteps.Form.RecordSource = RouteStepsSql(m_id)
        Me.sfLog.Form.RecordSource = RouteLogSql(m_id)
        m_sig = Signature(EditorValues())
    Else
        m_editable = False
        m_sig = ""
        Me.lblDraftRef.Caption = ""
    End If
    For Each c In Me.Controls
        If c.Name Like "tab_*" Or c.Name Like "ul_*" Or c.Name = "boxTabLine" Then SetVisible c, m_id <> 0
    Next
    SetVisible Me.lblDetailEmpty, m_id = 0
    ShowTab
End Sub

' Show the open tab's controls (none without a draft), then its state.
Private Sub ShowTab()
    Dim k As Variant
    If m_id <> 0 Then SelectTab Me, m_tab
    For Each k In Array("draft", "message", "route", "log")
        ShowGroup "grp" & UCase$(Left$(k, 1)) & Mid$(k, 2), m_id <> 0 And k = m_tab
    Next
    HideAlert Me, "alRoute"
    If m_id = 0 Then Exit Sub
    Select Case m_tab
        Case "draft": ShowDraftTab
        Case "message": ShowMessageTab
        Case "route": ShowRouteTab
    End Select
End Sub

' Show or hide the controls a tab's group lists (modBuild.GroupSince).
Private Sub ShowGroup(ByVal nm As String, ByVal isShown As Boolean)
    Dim v As Variant
    For Each v In Split(Me.Controls(nm).Caption, ";")
        If Len(v) > 0 Then SetVisible Me.Controls(v), isShown
    Next
End Sub

Private Sub ShowDraftTab()
    Dim v As Variant, why As String
    For Each v In Array("txtSubject", "cboType", "cboClass", "cboPrec", "cboClassif", "txtOriginator", "txtTo", "txtInfo", "txtParas")
        Me.Controls(v).Locked = Not m_editable
    Next
    SetVisible Me.btnSave, m_editable
    SetVisible Me.btnUndo, m_editable
    SetVisible Me.btnWithdraw, CanWithdraw()
    why = DraftEditProblem(m_id)
    Me.lblEditNote.Caption = IIf(Len(why) > 0, why, DraftStatusHint(m_d("status")))
    ShowMarking
End Sub

' The demo marking under the pickers, for a classification above UNCLASSIFIED.
Private Sub ShowMarking()
    If IsMarkedAboveUnclass(Me.cboClassif) Then
        Me.lblMarking.Caption = ClassificationMarking(Me.cboClassif) & " The text carries the test marking line."
    Else
        Me.lblMarking.Caption = ""
    End If
End Sub

Private Function CanWithdraw() As Boolean
    If m_d Is Nothing Then Exit Function
    If m_d("drafter_id") <> SignedInPersonID() Then Exit Function
    CanWithdraw = (m_d("status") = DS_DRAFTING Or m_d("status") = DS_RETURNED Or m_d("status") = DS_ROUTING)
End Function

Private Sub ShowMessageTab()
    If m_d("status") = DS_RELEASED Then
        Me.txtPreview = m_d("rendered")
        Me.lblPreviewNote.Caption = "Released " & m_d("dtg") & " as " & m_d("local_seq") & ": the text is frozen as transmitted."
    Else
        Me.txtPreview = DraftPreviewText(EditorValues())
        Me.lblPreviewNote.Caption = "The DTG and sequence are filled in at release. The check uses the ones it would take now."
    End If
    RunCheck
End Sub

' The naval format check: the released text as it went, or the draft as it
' would go if released now. Blocking issues first, then the warnings.
Private Sub RunCheck()
    Dim issues As String, blocking As String, warns As String
    If m_d("status") = DS_RELEASED Then
        issues = NavalFormatIssues(m_d("rendered"))
    Else
        issues = DraftFormatIssues(EditorValues())
    End If
    blocking = FormatIssuesOf(issues, False)
    warns = FormatIssuesOf(issues, True)
    If Len(issues) = 0 Then
        Me.txtIssues = "The check found nothing to fix."
    Else
        Me.txtIssues = IIf(Len(blocking) > 0, "To fix before it can be submitted or released:" & vbCrLf & blocking, "Nothing blocks it.") & _
            IIf(Len(warns) > 0, vbCrLf & "Warnings (advisory):" & vbCrLf & warns, "")
    End If
End Sub

' Where the draft stands and what the signed-in person may do with it.
Private Sub ShowRouteTab()
    Dim st As String, wf As Long, why As String, stepId As Long, holder As Long, title As String, action As String
    Dim mine As Boolean, isRelease As Boolean
    st = m_d("status")
    SetVisible Me.btnSubmit, False
    SetVisible Me.btnEndorse, False
    SetVisible Me.btnRelease, False
    SetVisible Me.btnReturn, False
    SetVisible Me.btnOpenMessage, False
    SetVisible Me.capComment, False
    SetVisible Me.txtComment, False
    Select Case st
        Case DS_DRAFTING, DS_RETURNED
            Me.lblRouteState.Caption = IIf(st = DS_RETURNED, "Returned for correction", "Not yet submitted")
            wf = WorkflowFor(m_d("msg_type"), m_d("originator"))
            If wf = 0 Then
                Me.lblRouteInfo.Caption = "No routing workflow matches " & m_d("msg_type") & " from " & m_d("originator") & " yet."
            Else
                Me.lblRouteInfo.Caption = "Submitting sends it on " & WorkflowName(wf) & ": " & WorkflowChainText(wf) & "."
            End If
            If m_d("drafter_id") = SignedInPersonID() Then
                SetVisible Me.btnSubmit, True
                Me.btnSubmit.Caption = IIf(st = DS_RETURNED, "Resubmit", "Submit for routing")
                why = SubmitProblem(m_id, EditorValues())
                If Len(why) > 0 Then ShowAlert Me, "alRoute", ClipText(Replace(why, vbCrLf, " "), 220), "warning"
            Else
                ShowAlert Me, "alRoute", "Only its drafter, " & m_d("drafter_name") & ", submits it.", "info"
            End If
        Case DS_ROUTING
            stepId = DraftCurrentStepID(m_id)
            holder = Nz(DLookup("PersonID", "tblDraftRoute", "RouteStepID=" & stepId), 0)
            title = Nz(DLookup("StepTitle", "tblDraftRoute", "RouteStepID=" & stepId), "")
            action = Nz(DLookup("StepAction", "tblDraftRoute", "RouteStepID=" & stepId), "review")
            isRelease = (action = "release")
            mine = (holder <> 0 And holder = SignedInPersonID())
            Me.lblRouteState.Caption = "In routing: " & title & IIf(isRelease, " (release)", "")
            Me.lblRouteInfo.Caption = "On " & WorkflowName(m_d("workflow_id")) & ". It waits on " & CurrentHolderText() & "."
            If mine Then
                SetVisible Me.capComment, True
                SetVisible Me.txtComment, True
                SetVisible Me.btnReturn, True
                If isRelease Then
                    SetVisible Me.btnRelease, True
                    why = ReleaseProblem(m_id)
                    If Len(why) > 0 Then
                        ShowAlert Me, "alRoute", why, "warning"
                    Else
                        ShowAlert Me, "alRoute", "You hold the release step. Release assigns the DTG and the next sequence, " & _
                            LocalSequenceText(SerialText(NextSequenceNumber())) & ".", "info"
                    End If
                Else
                    SetVisible Me.btnEndorse, True
                    ShowAlert Me, "alRoute", "It waits on you. Endorse passes it to the next step; Return sends it back with your comment.", "info"
                End If
            Else
                why = HolderProblem(holder, Nz(DLookup("PersonName", "tblDraftRoute", "RouteStepID=" & stepId), ""))
                If Len(why) > 0 Then ShowAlert Me, "alRoute", why, "warning"
            End If
        Case DS_RELEASED
            Me.lblRouteState.Caption = "Released " & m_d("dtg") & " as " & m_d("local_seq")
            Me.lblRouteInfo.Caption = "Released by " & m_d("released_by") & IIf(Len(m_d("designation")) > 0, " as " & m_d("designation"), "") & _
                " on " & Format$(m_d("released_on"), "d mmm yyyy hh:nn") & ". The record is frozen."
            SetVisible Me.btnOpenMessage, m_d("message_id") <> 0
        Case Else
            Me.lblRouteState.Caption = "Withdrawn"
            Me.lblRouteInfo.Caption = DraftStatusHint(st)
    End Select
End Sub

' "LT J. Taylor (N3 review)" for the current step, "nobody" without one.
Private Function CurrentHolderText() As String
    Dim stepId As Long
    stepId = DraftCurrentStepID(m_id)
    If stepId = 0 Then CurrentHolderText = "nobody": Exit Function
    CurrentHolderText = Nz(DLookup("PersonName", "tblDraftRoute", "RouteStepID=" & stepId), "") & " (" & _
        Nz(DLookup("StepTitle", "tblDraftRoute", "RouteStepID=" & stepId), "") & ")"
End Function

' ---------------------------------------------------------------------------
' Editing
' ---------------------------------------------------------------------------

' The editor's values (NewDraftValues' keys).
Private Function EditorValues() As Object
    Dim d As Object
    Set d = NewDraftValues()
    d("subject") = Trim$(Nz(Me.txtSubject, ""))
    d("msg_type") = Nz(Me.cboType, "")
    d("msg_class") = Nz(Me.cboClass, "")
    d("precedence") = Nz(Me.cboPrec, "")
    d("classification") = Nz(Me.cboClassif, "")
    d("originator") = UCase$(Trim$(Nz(Me.txtOriginator, "")))
    d("to_line") = Trim$(Nz(Me.txtTo, ""))
    d("info_line") = Trim$(Nz(Me.txtInfo, ""))
    d("paragraphs") = Nz(Me.txtParas, "")
    Set EditorValues = d
End Function

Private Function Signature(ByVal d As Object) As String
    Dim k As Variant
    For Each k In d.Keys
        Signature = Signature & k & "=" & d(k) & vbTab
    Next
End Function

' True when the shown draft has edits not saved.
Private Function DirtyNow() As Boolean
    If m_id = 0 Or Not m_editable Or Len(m_sig) = 0 Then Exit Function
    DirtyNow = (Signature(EditorValues()) <> m_sig)
End Function

' Save the editor's values. Returns True when saved.
Private Function SaveClicked() As Boolean
    Dim why As String
    why = SaveDraft(m_id, EditorValues())
    If Len(why) > 0 Then KitMsgBox why, vbExclamation, MSG_TITLE: Exit Function
    KitToast "Draft saved.", "success"
    RefreshList m_id
    SaveClicked = True
End Function

' Before the page or the shown draft changes: True to go on (saved, or
' nothing to save, or the edits dropped), False to stay.
Private Function OfferSave() As Boolean
    If Not DirtyNow() Then OfferSave = True: Exit Function
    Select Case KitMsgBox("Save your changes to this draft?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: OfferSave = SaveClicked()
        Case vbNo: m_sig = "": OfferSave = True
        Case Else: OfferSave = False
    End Select
End Function
