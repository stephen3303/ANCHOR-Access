Option Compare Database
Option Explicit

' Subscriptions (see modBuildSubscriptions.Build_pgSubscriptions). The
' catalogue (tblSubscriptions, active ones), mandatory first, with the
' signed-in person's status of each (modSubscriptions.SubsListSql),
' filtered by the search box and the category. The picked one shows on the
' right: its category, default delivery and description, where the person
' stands (alStatus), what it matches and how many messages match now
' (modSubscriptions.SubCriteriaText, SubMatchCount), and its actions:
'   Request...        not subscribed, or denied: dlgRequestSub (needs the
'                     Subscription Requester role, checked again at the write)
'   Pause / Resume    their own Active or Paused, non-mandatory subscription
'   Withdraw request  their own pending request
' A mandatory subscription always shows Active, assigned by command, with
' no action. With no data file linked the page says so.

Private Const MSG_TITLE As String = "Subscriptions"
Private Const ALL_CATS As String = "(All categories)"

Private m_search As String
Private m_cat As String         ' "" for every category
Private m_id As Long            ' the subscription shown on the right (0: none)
Private m_psId As Long          ' the person's latest tblPersonSubs row for it (0: none)
Private m_status As String      ' their status of it

Public Property Get HelpTopic() As String
    HelpTopic = "subscriptions"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.cboCategory.RowSource = SubCategoriesList()
    Me.cboCategory = ALL_CATS
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no subscriptions."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        SetEnabled Me.cboCategory, False
        ShowSub 0
        Exit Sub
    End If
    RefreshList 0
    Exit Sub
Err_Handler:
    HandleError "pgSubscriptions.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' Search and category (the kit's search box calls SearchFor)
' ---------------------------------------------------------------------------

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList m_id
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgSubscriptions.SearchFor", text
End Sub

Private Sub cboCategory_AfterUpdate()
    On Error GoTo Err_Handler
    If IsNull(Me.cboCategory) Then Me.cboCategory = ALL_CATS
    m_cat = IIf(Me.cboCategory = ALL_CATS, "", Nz(Me.cboCategory, ""))
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgSubscriptions.cboCategory_AfterUpdate"
End Sub

' The empty state's button: every category, no search.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    m_cat = ""
    Me.cboCategory = ALL_CATS
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgSubscriptions.emptyList_b_Click"
End Sub

' ---------------------------------------------------------------------------
' The list and the detail
' ---------------------------------------------------------------------------

' Load the catalogue for the category and the search, and show openId (0,
' or not listed: the first).
Private Sub RefreshList(ByVal openId As Long)
    Dim n As Long, rs As DAO.Recordset
    If Not DataLinked() Then Exit Sub
    Me.sfList.Form.RecordSource = SubsListSql(m_cat, m_search)
    GridHighlight Me.sfList, m_search
    Set rs = Me.sfList.Form.RecordsetClone
    If rs.RecordCount > 0 Then rs.MoveLast
    n = rs.RecordCount
    ShowEmpty Me, "emptyList", n = 0
    SetVisible Me.emptyList_b, n = 0 And (Len(m_cat) > 0 Or Len(m_search) > 0)
    SetVisible Me.sfList, n > 0
    If GridSelect(Me.sfList, openId) Then
        ShowSub openId
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ShowSub 0
    End If
End Sub

' A row clicked (the table calls this): show it on the right.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    If grid <> "sfList" Then Exit Sub
    ShowSub Nz(Me.sfList.Form.txtSel, 0)
    Exit Sub
Err_Handler:
    HandleError "pgSubscriptions.GridClick", grid
End Sub

' Show subscription subId on the right (0: the hint instead).
Private Sub ShowSub(ByVal subId As Long)
    Dim rs As DAO.Recordset, isOpen As Boolean, mand As Boolean, c As Control, n As Long, delivery As String
    m_id = subId: m_psId = 0: m_status = ""
    isOpen = (subId <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT SubName, Category, Description, IsMandatory, DefaultFrequency FROM tblSubscriptions " & _
            "WHERE SubID=" & subId, dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_id = 0
        Else
            mand = Nz(rs!IsMandatory, False)
            Me.lblSubName.Caption = Replace(rs!SubName, "&", "&&")
            Me.lblSubMeta.Caption = Replace(Nz(rs!Category, "No category"), "&", "&&") & ". Delivered " & _
                Nz(rs!DefaultFrequency, "Real-time") & " by default."
            Me.lblSubDesc.Caption = Replace(Nz(rs!Description, ""), "&", "&&")
            m_status = MySubStatus(subId)
            If Not mand Then m_psId = MyPersonSubID(subId)
            If m_psId <> 0 Then delivery = SubDeliveryText(m_psId)
            ShowStatus mand, delivery
            Me.txtCriteria = SubCriteriaText(subId)
            n = SubMatchCount(subId)
            Me.lblMatches.Caption = n & " " & Plural(n, "message") & " in ANCHOR match it now."
            SetEnabled Me.btnRequest, Not mand And (m_status = SUB_NOT_SUBSCRIBED Or m_status = "Denied") And _
                (HasRole("REQUEST") Or IsAnchorAdmin())
            SetEnabled Me.btnPause, Not mand And m_psId <> 0 And (m_status = "Active" Or m_status = "Paused")
            Me.btnPause.Caption = IIf(m_status = "Paused", "Resume", "Pause")
            SetEnabled Me.btnWithdraw, Not mand And m_psId <> 0 And m_status = "Pending Approval"
        End If
        rs.Close
    End If
    For Each c In Me.Controls
        Select Case True
            Case c.Name = "lblSubName", c.Name = "lblSubMeta", c.Name = "lblSubDesc", c.Name = "capCriteria", c.Name = "txtCriteria", _
                    c.Name = "capMatches", c.Name = "lblMatches"
                SetVisible c, isOpen
            ' A mandatory subscription can't be requested, paused or
            ' withdrawn: its buttons are hidden, not just greyed, and the
            ' banner says why.
            Case c.Name = "btnRequest", c.Name = "btnPause", c.Name = "btnWithdraw"
                SetVisible c, isOpen And Not mand
        End Select
    Next
    If Not isOpen Then HideAlert Me, "alStatus"     ' ShowAlert (ShowStatus) shows it
    SetVisible Me.lblDetailEmpty, Not isOpen
End Sub

' Where the person stands with the shown subscription, in alStatus.
Private Sub ShowStatus(ByVal mand As Boolean, ByVal delivery As String)
    Dim rs As DAO.Recordset, by As String, msg As String
    If mand Then
        ShowAlert Me, "alStatus", "Mandatory: you receive it whatever you choose, and it can't be paused or withdrawn.", "success", _
            "Active - Assigned by command"
        Exit Sub
    End If
    If m_psId <> 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT RequestedOn, DeciderName, DecidedOn, DecisionNote FROM tblPersonSubs WHERE PersonSubID=" & _
            m_psId, dbOpenSnapshot)
        If Not rs.EOF Then
            by = Nz(rs!DeciderName, "")
            If Not IsNull(rs!DecidedOn) Then by = IIf(Len(by) > 0, by, "an approver") & " on " & Format$(rs!DecidedOn, "dd mmm yyyy")
            Select Case m_status
                Case "Pending Approval"
                    msg = "Requested" & IIf(IsNull(rs!RequestedOn), "", " on " & Format$(rs!RequestedOn, "dd mmm yyyy")) & _
                        ". Waiting for your unit's Subscription Approver."
                Case "Denied"
                    msg = "Denied" & IIf(Len(by) > 0, " by " & by, "") & IIf(IsNull(rs!DecisionNote), ".", ": " & rs!DecisionNote & ".") & _
                        " You can request it again."
                Case Else
                    msg = IIf(Len(by) > 0, "Approved by " & by & ". ", "")
            End Select
        End If
        rs.Close
    End If
    Select Case m_status
        Case "Active"
            ShowAlert Me, "alStatus", msg & "You receive it: " & delivery & ". Pause stops it for a while.", "success", "Active"
        Case "Paused"
            ShowAlert Me, "alStatus", msg & "You don't receive it while it is paused. Resume starts it again (" & delivery & ").", "info", "Paused"
        Case "Pending Approval"
            ShowAlert Me, "alStatus", msg, "warning", "Pending Approval"
        Case "Denied"
            ShowAlert Me, "alStatus", msg, "destructive", "Denied"
        Case Else
            ShowAlert Me, "alStatus", "Request it to start receiving it. Your unit's Subscription Approver decides.", "info", "Not subscribed"
    End Select
End Sub

' ---------------------------------------------------------------------------
' Actions
' ---------------------------------------------------------------------------

Private Sub btnRequest_Click()
    Dim nm As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    If Not DataLinked() Then Exit Sub
    nm = Nz(DLookup("SubName", "tblSubscriptions", "SubID=" & m_id), "")
    If RunDialog("dlgRequestSub", "sub=" & m_id) Then
        KitToast "Request for " & ClipText(nm, 60) & " sent for approval.", "success"
        RefreshList m_id
    End If
    Exit Sub
Err_Handler:
    HandleError "pgSubscriptions.btnRequest_Click", m_id
End Sub

Private Sub btnPause_Click()
    Dim problem As String, pausing As Boolean
    On Error GoTo Err_Handler
    If m_psId = 0 Then Exit Sub
    pausing = (m_status = "Active")
    problem = SetSubscriptionPaused(m_psId, pausing)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbInformation, MSG_TITLE
    Else
        KitToast IIf(pausing, "Subscription paused.", "Subscription resumed."), "success"
    End If
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgSubscriptions.btnPause_Click", m_psId
End Sub

Private Sub btnWithdraw_Click()
    Dim problem As String
    On Error GoTo Err_Handler
    If m_psId = 0 Then Exit Sub
    If KitMsgBox("Withdraw your request for " & ClipText(Nz(DLookup("SubName", "tblSubscriptions", "SubID=" & m_id), ""), 60) & "?", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
    problem = WithdrawSubRequest(m_psId)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE Else KitToast "Request withdrawn.", "success"
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgSubscriptions.btnWithdraw_Click", m_psId
End Sub
