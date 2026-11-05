Option Compare Database
Option Explicit

' Admin: Routing (see modBuildDrafting.Build_pgRouting), reached from the
' Admin page. Three tabs:
'   Workflows          the routing workflows (modRouting.WorkflowListSql) and
'                      the picked one: its name, description and what it
'                      matches (Save, Undo), turned on or off, and its chain,
'                      whose steps change at once (dlgRouteStep, Up, Down,
'                      Remove); New workflow adds one, off
'   Release authority  who may release (AuthorityListSql), Designate and
'                      Remove designation
'   Command settings   the text case, FM line, MSGID originator and the next
'                      sequence number (Save, Undo)
' Every change is checked again in the data file (RequireAnchorAdmin) and
' audited (modRouting). The page is for ANCHOR admins (modSignIn.PageAllowed).
'
' Owner: drafting and release. Depends on: modErrors, modRouting, modSignIn,
' modUI; the kit's modKitUI, modKitUtil.

Private Const MSG_TITLE As String = "Routing"

Private m_tab As String         ' wf, ra or cs
Private m_wf As Long            ' the workflow shown (0: none)
Private m_wfOn As Boolean       ' it is on
Private m_wfSig As String       ' its name, description and matches as loaded
Private m_csSig As String       ' the settings as loaded
Private m_nWf As Long           ' how many workflows

Public Property Get HelpTopic() As String
    HelpTopic = "routing"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_tab = "wf"
    If Not DataLinked() Then
        ShowEmpty Me, "emptyWf", True, "No data file linked", "This copy was built without the data file, so there are no workflows."
        SetVisible Me.sfWf, False
        ShowTab
        SetEnabled Me.btnNewWf, False
        ShowWf 0
        SetEnabled Me.btnAddRa, False
        SetEnabled Me.btnCsSave, False
        Exit Sub
    End If
    Me.cboRaPerson.RowSource = RosterSql()
    Me.cboRaScope.RowSource = ScopeList()
    Me.cboRaScope = "*"
    RefreshWf 0
    RefreshRa
    LoadSettings
    ShowTab
    Exit Sub
Err_Handler:
    HandleError "pgRouting.Form_Load"
End Sub

Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = OfferSaveWf()
    If CanLeave Then CanLeave = OfferSaveCs()
    Exit Function
Err_Handler:
    HandleError "pgRouting.CanLeave"
    CanLeave = True
End Function

Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            If m_tab = "wf" And m_wf <> 0 Then SaveWf
            If m_tab = "cs" Then SaveCs
            Shortcut = True
        Case SHORTCUT_NEW
            If m_tab = "wf" And Me.btnNewWf.Enabled Then btnNewWf_Click
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgRouting.Shortcut", action
    Shortcut = True
End Function

' The value list of the Message types picker: every type, or one ("*" for all).
Private Function ScopeList() As String
    Dim rs As DAO.Recordset, s As String
    s = "*;""(All message types)"""
    Set rs = CurrentDb.OpenRecordset("SELECT TypeCode, TypeName FROM tblMessageTypes WHERE IsActive ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & ";""" & Replace(rs!TypeCode, """", """""") & """;""" & Replace(rs!TypeCode & " - " & Nz(rs!TypeName, ""), """", """""") & """"
        rs.MoveNext
    Loop
    rs.Close
    ScopeList = s
End Function

' ---------------------------------------------------------------------------
' Tabs
' ---------------------------------------------------------------------------

Public Function TabPick(ByVal key As String)
    On Error GoTo Err_Handler
    If key = m_tab Then Exit Function
    If m_tab = "wf" Then
        If Not OfferSaveWf() Then Exit Function
    ElseIf m_tab = "cs" Then
        If Not OfferSaveCs() Then Exit Function
    End If
    m_tab = key
    ShowTab
    Exit Function
Err_Handler:
    HandleError "pgRouting.TabPick", key
End Function

Private Sub ShowTab()
    SelectTab Me, m_tab
    ShowGroup "grpWf", m_tab = "wf"
    ShowGroup "grpRa", m_tab = "ra"
    ShowGroup "grpCs", m_tab = "cs"
    If m_tab = "wf" Then
        ShowEmpty Me, "emptyWf", m_nWf = 0 Or Not DataLinked()
        SetVisible Me.sfWf, m_nWf > 0
        ShowWf m_wf
    Else
        ShowEmpty Me, "emptyWf", False
    End If
End Sub

Private Sub ShowGroup(ByVal nm As String, ByVal isShown As Boolean)
    Dim v As Variant
    For Each v In Split(Me.Controls(nm).Caption, ";")
        If Len(v) > 0 Then SetVisible Me.Controls(v), isShown
    Next
End Sub

Private Sub btnBack_Click()
    On Error GoTo Err_Handler
    NavTo "admin"
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnBack_Click"
End Sub

' ---------------------------------------------------------------------------
' Rows
' ---------------------------------------------------------------------------

Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    Select Case grid
        Case "sfWf"
            id = Nz(Me.sfWf.Form.txtSel, 0)
            If id = m_wf Then Exit Sub
            If Not OfferSaveWf() Then GridSelect Me.sfWf, m_wf: Exit Sub
            ShowWf id
        Case "sfWfSteps"
            If dbl Then EditStep Nz(Me.sfWfSteps.Form.txtSel, 0)
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgRouting.GridClick", grid & " " & id
End Sub

' ---------------------------------------------------------------------------
' Workflows
' ---------------------------------------------------------------------------

Private Sub RefreshWf(ByVal openId As Long)
    Me.sfWf.Form.RecordSource = WorkflowListSql()
    m_nWf = Me.sfWf.Form.RecordsetClone.RecordCount
    If m_tab = "wf" Then
        ShowEmpty Me, "emptyWf", m_nWf = 0
        SetVisible Me.sfWf, m_nWf > 0
    End If
    If GridSelect(Me.sfWf, openId) Then
        ShowWf openId
    ElseIf IsNull(SelectFirstRow(Me.sfWf)) Then
        ShowWf 0
    Else
        ShowWf Nz(Me.sfWf.Form.txtSel, 0)
    End If
End Sub

' Show workflow wfId in the editor (0: an empty, disabled editor).
Private Sub ShowWf(ByVal wfId As Long)
    Dim rs As DAO.Recordset, v As Variant, why As String
    m_wf = 0: m_wfOn = False
    If wfId <> 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblWorkflows WHERE WorkflowID=" & wfId, dbOpenSnapshot)
        If Not rs.EOF Then
            m_wf = wfId
            m_wfOn = Nz(rs!IsActive, False)
            Me.txtWfName = rs!WorkflowName
            Me.txtWfDesc = rs!Description
            Me.txtWfTypes = rs!MatchTypes
            Me.txtWfOrigs = rs!MatchOriginators
        End If
        rs.Close
    End If
    If m_wf = 0 Then
        Me.txtWfName = Null: Me.txtWfDesc = Null: Me.txtWfTypes = Null: Me.txtWfOrigs = Null
        If DataLinked() Then Me.sfWfSteps.Form.RecordSource = WorkflowStepsSql(0)
        Me.lblWfState.Caption = ""
        Me.lblWfNote.Caption = IIf(DataLinked(), "Pick a workflow in the list, or add one with New workflow.", "")
    Else
        Me.sfWfSteps.Form.RecordSource = WorkflowStepsSql(m_wf)
        Me.lblWfState.Caption = IIf(m_wfOn, "On: drafts that match it take it.", "Off: no draft takes it.")
        why = WorkflowProblem(m_wf)
        If Len(why) = 0 Then why = StrandedText(m_wf)
        Me.lblWfNote.Caption = IIf(Len(why) > 0, why, "Chain: " & WorkflowChainText(m_wf) & ". Every step is held.") & _
            " Changes to the chain apply to drafts submitted from now on; drafts in routing keep their route."
    End If
    Me.btnWfOnOff.Caption = IIf(m_wfOn, "Turn off", "Turn on")
    For Each v In Array("txtWfName", "txtWfDesc", "txtWfTypes", "txtWfOrigs")
        Me.Controls(v).Locked = (m_wf = 0)
    Next
    For Each v In Array("btnWfOnOff", "btnAddStep", "btnEditStep", "btnStepUp", "btnStepDown", "btnRemoveStep", "btnWfSave", "btnWfUndo")
        SetEnabled Me.Controls(v), m_wf <> 0
    Next
    m_wfSig = WfSignature()
End Sub

Private Function WfSignature() As String
    WfSignature = Nz(Me.txtWfName, "") & vbTab & Nz(Me.txtWfDesc, "") & vbTab & Nz(Me.txtWfTypes, "") & vbTab & Nz(Me.txtWfOrigs, "")
End Function

Private Function WfDirty() As Boolean
    If m_wf = 0 Then Exit Function
    WfDirty = (WfSignature() <> m_wfSig)
End Function

Private Function SaveWf() As Boolean
    Dim why As String
    why = SaveWorkflow(m_wf, Nz(Me.txtWfName, ""), Nz(Me.txtWfDesc, ""), Nz(Me.txtWfTypes, ""), Nz(Me.txtWfOrigs, ""))
    If Len(why) > 0 Then KitMsgBox why, vbExclamation, MSG_TITLE: Exit Function
    KitToast "Workflow saved.", "success"
    RefreshWf m_wf
    SaveWf = True
End Function

Private Function OfferSaveWf() As Boolean
    If Not WfDirty() Then OfferSaveWf = True: Exit Function
    Select Case KitMsgBox("Save your changes to this workflow?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: OfferSaveWf = SaveWf()
        Case vbNo: m_wfSig = WfSignature(): OfferSaveWf = True
        Case Else: OfferSaveWf = False
    End Select
End Function

Private Sub btnNewWf_Click()
    Dim id As Long
    On Error GoTo Err_Handler
    If Not OfferSaveWf() Then Exit Sub
    id = NewWorkflow()
    RefreshWf id
    KitToast "Workflow added, off. Name it, say what it matches, pick who holds each step, then turn it on.", "success"
    Me.txtWfName.SetFocus
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnNewWf_Click"
End Sub

Private Sub btnWfSave_Click()
    On Error GoTo Err_Handler
    SaveWf
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnWfSave_Click", m_wf
End Sub

Private Sub btnWfUndo_Click()
    On Error GoTo Err_Handler
    ShowWf m_wf
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnWfUndo_Click", m_wf
End Sub

Private Sub btnWfOnOff_Click()
    Dim why As String
    On Error GoTo Err_Handler
    If m_wf = 0 Then Exit Sub
    If Not OfferSaveWf() Then Exit Sub
    why = SetWorkflowActive(m_wf, Not m_wfOn)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: Exit Sub
    KitToast IIf(m_wfOn, "Workflow turned off.", "Workflow turned on."), "success"
    RefreshWf m_wf
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnWfOnOff_Click", m_wf
End Sub

Private Sub btnAddStep_Click()
    On Error GoTo Err_Handler
    EditStep 0
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnAddStep_Click", m_wf
End Sub

Private Sub btnEditStep_Click()
    Dim stepId As Long
    On Error GoTo Err_Handler
    stepId = Nz(Me.sfWfSteps.Form.txtSel, 0)
    If stepId = 0 Then KitMsgBox "Pick a step in the chain first.", vbInformation, MSG_TITLE: Exit Sub
    EditStep stepId
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnEditStep_Click", m_wf
End Sub

' Add a step (stepId 0) or change one, in dlgRouteStep.
Private Sub EditStep(ByVal stepId As Long)
    If m_wf = 0 Then Exit Sub
    If RunDialog("dlgRouteStep", "wf=" & m_wf & ";step=" & stepId) Then
        KitToast IIf(stepId = 0, "Step added.", "Step saved."), "success"
        RefreshSteps stepId
    End If
End Sub

' Reload the chain and the note under it, keeping the workflow's unsaved
' edits, and pick step stepId (0: the last).
Private Sub RefreshSteps(ByVal stepId As Long)
    Dim sig As String, nm As Variant, desc As Variant, types As Variant, origs As Variant, why As String
    sig = m_wfSig
    nm = Me.txtWfName: desc = Me.txtWfDesc: types = Me.txtWfTypes: origs = Me.txtWfOrigs
    If stepId = 0 Then
        stepId = Nz(DLookup("WfStepID", "tblWorkflowSteps", "WorkflowID=" & m_wf & " AND StepOrder=" & _
            Nz(DMax("StepOrder", "tblWorkflowSteps", "WorkflowID=" & m_wf), 0)), 0)
    End If
    GridSelect Me.sfWfSteps, stepId, True
    why = WorkflowProblem(m_wf)
    If Len(why) = 0 Then why = StrandedText(m_wf)
    Me.lblWfNote.Caption = IIf(Len(why) > 0, why, "Chain: " & WorkflowChainText(m_wf) & ". Every step is held.") & _
        " Changes to the chain apply to drafts submitted from now on; drafts in routing keep their route."
    Me.txtWfName = nm: Me.txtWfDesc = desc: Me.txtWfTypes = types: Me.txtWfOrigs = origs
    m_wfSig = sig
End Sub

Private Sub btnStepUp_Click()
    On Error GoTo Err_Handler
    MoveStep -1
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnStepUp_Click", m_wf
End Sub

Private Sub btnStepDown_Click()
    On Error GoTo Err_Handler
    MoveStep 1
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnStepDown_Click", m_wf
End Sub

Private Sub MoveStep(ByVal delta As Long)
    Dim stepId As Long, why As String
    stepId = Nz(Me.sfWfSteps.Form.txtSel, 0)
    If stepId = 0 Then KitMsgBox "Pick a step in the chain first.", vbInformation, MSG_TITLE: Exit Sub
    why = MoveWorkflowStep(m_wf, stepId, delta)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: Exit Sub
    RefreshSteps stepId
End Sub

Private Sub btnRemoveStep_Click()
    Dim stepId As Long, why As String
    On Error GoTo Err_Handler
    stepId = Nz(Me.sfWfSteps.Form.txtSel, 0)
    If stepId = 0 Then KitMsgBox "Pick a step in the chain first.", vbInformation, MSG_TITLE: Exit Sub
    If KitMsgBox("Take " & Nz(DLookup("StepTitle", "tblWorkflowSteps", "WfStepID=" & stepId), "this step") & " off the chain?", _
            vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
    why = RemoveWorkflowStep(m_wf, stepId)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: Exit Sub
    KitToast "Step removed.", "success"
    RefreshSteps 0
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnRemoveStep_Click", m_wf
End Sub

' ---------------------------------------------------------------------------
' Release authority
' ---------------------------------------------------------------------------

Private Sub RefreshRa()
    Me.sfRa.Form.RecordSource = AuthorityListSql()
    SelectFirstRow Me.sfRa
End Sub

Private Sub btnAddRa_Click()
    Dim why As String, scopeType As String, pid As Long
    On Error GoTo Err_Handler
    scopeType = Nz(Me.cboRaScope, "*")
    If scopeType = "*" Then scopeType = ""
    pid = Nz(Me.cboRaPerson, 0)
    why = AddReleaseAuthority(pid, Nz(Me.txtRaDesig, ""), Nz(Me.chkRaCO, False), scopeType)
    If Len(why) > 0 Then KitMsgBox why, vbExclamation, MSG_TITLE: Exit Sub
    If PersonHasRole(pid, "RELEASE") Then
        KitToast RosterName(pid) & " is a release authority.", "success"
    Else
        KitToast RosterName(pid) & " is a release authority. Give them the Designated Release Authority role under Admin: Users too.", "success"
    End If
    Me.cboRaPerson = Null: Me.txtRaDesig = Null: Me.chkRaCO = False: Me.cboRaScope = "*"
    RefreshRa
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnAddRa_Click"
End Sub

Private Sub btnRemoveRa_Click()
    Dim aid As Long
    On Error GoTo Err_Handler
    aid = Nz(Me.sfRa.Form.txtSel, 0)
    If aid = 0 Then KitMsgBox "Pick a designation in the list first.", vbInformation, MSG_TITLE: Exit Sub
    If KitMsgBox("Take away " & RosterName(DLookup("PersonID", "tblReleaseAuthorities", "AuthorityID=" & aid)) & "'s designation as " & _
            ReleaseDesignationOf(aid) & "? They can no longer release under it.", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
    RemoveReleaseAuthority aid
    KitToast "Designation removed.", "success"
    RefreshRa
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnRemoveRa_Click", aid
End Sub

' ---------------------------------------------------------------------------
' Command settings
' ---------------------------------------------------------------------------

Private Sub LoadSettings()
    Me.cboCsCase = CommandTextCase()
    Me.txtCsOrig = CommandOriginator()
    Me.txtCsMsgId = MsgIdOriginator()
    Me.txtCsNext = NextSequenceNumber()
    Me.lblCsSeq.Caption = "The local sequence is # and four digits (NAVEDTRA 14226), as the naval format check requires, so " & _
        "the prefix and digits don't change here. The next release takes " & LocalSequenceText(SerialText(NextSequenceNumber())) & _
        ", and every release takes the next number in turn."
    m_csSig = CsSignature()
End Sub

Private Function CsSignature() As String
    CsSignature = Nz(Me.cboCsCase, "") & vbTab & Nz(Me.txtCsOrig, "") & vbTab & Nz(Me.txtCsMsgId, "") & vbTab & Nz(Me.txtCsNext, "")
End Function

Private Function SaveCs() As Boolean
    Dim why As String
    why = SaveCommandSettings(Nz(Me.cboCsCase, ""), Nz(Me.txtCsNext, ""), Nz(Me.txtCsOrig, ""), Nz(Me.txtCsMsgId, ""))
    If Len(why) > 0 Then KitMsgBox why, vbExclamation, MSG_TITLE: Exit Function
    KitToast "Settings saved.", "success"
    LoadSettings
    SaveCs = True
End Function

Private Function OfferSaveCs() As Boolean
    If Not DataLinked() Or CsSignature() = m_csSig Then OfferSaveCs = True: Exit Function
    Select Case KitMsgBox("Save your changes to the command settings?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: OfferSaveCs = SaveCs()
        Case vbNo: m_csSig = CsSignature(): OfferSaveCs = True
        Case Else: OfferSaveCs = False
    End Select
End Function

Private Sub btnCsSave_Click()
    On Error GoTo Err_Handler
    SaveCs
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnCsSave_Click"
End Sub

Private Sub btnCsUndo_Click()
    On Error GoTo Err_Handler
    LoadSettings
    Exit Sub
Err_Handler:
    HandleError "pgRouting.btnCsUndo_Click"
End Sub
