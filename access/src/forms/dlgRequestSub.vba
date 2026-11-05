Option Compare Database
Option Explicit

' Request a subscription (modBuildSubscriptions.Build_dlgRequestSub), from
' the Subscriptions page. OpenArgs "sub=<SubID>". The person picks how often
' (Real-time, Daily Digest or Weekly; the subscription's default first),
' the channels (In app, Email, Teams; In app ticked) and says why, which is
' required. Submit for approval calls modSubscriptions.RequestSubscription,
' which checks the Subscription Requester role in the data file and adds a
' tblPersonSubs row, Status 'Pending Approval'; RunDialog then sees
' DialogResult "ok". Cancel, or Esc, changes nothing.

Private Const MSG_TITLE As String = "Request a subscription"

Private m_subId As Long

Public Property Get HelpTopic() As String
    HelpTopic = "subscriptions.request"
End Property

Private Sub Form_Load()
    Dim nm As String, freq As String
    On Error GoTo Err_Handler
    m_subId = OpenArgLong(Me, "sub")
    nm = Nz(DLookup("SubName", "tblSubscriptions", "SubID=" & m_subId), "")
    freq = Nz(DLookup("DefaultFrequency", "tblSubscriptions", "SubID=" & m_subId), "")
    Me.lblHead.Caption = "Request " & Replace(ClipText(nm, 60), "&", "&&")
    Me.lblSub.Caption = "Your request goes to your unit's Subscription Approver. You start receiving it once it is approved."
    Select Case freq
        Case "Daily Digest": SetSegment Me, "segFreq", "daily"
        Case "Weekly": SetSegment Me, "segFreq", "weekly"
        Case Else: SetSegment Me, "segFreq", "rt"
    End Select
    Me.chkInApp = True
    Me.chkEmail = False
    Me.chkTeams = False
    CenterOnApp Me
    Me.txtWhy.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgRequestSub.Form_Load", Nz(Me.OpenArgs, "")
End Sub

Private Sub btnSubmit_Click()
    Dim problem As String, freq As String
    On Error GoTo Err_Handler
    Select Case SegmentValue(Me, "segFreq")
        Case "daily": freq = "Daily Digest"
        Case "weekly": freq = "Weekly"
        Case Else: freq = "Real-time"
    End Select
    problem = RequestSubscription(m_subId, freq, Nz(Me.chkInApp, False), Nz(Me.chkEmail, False), Nz(Me.chkTeams, False), _
        FieldText(Me, Me.txtWhy))
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, MSG_TITLE
        Me.txtWhy.SetFocus
        Exit Sub
    End If
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgRequestSub.btnSubmit_Click", m_subId
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgRequestSub.btnCancel_Click"
End Sub
