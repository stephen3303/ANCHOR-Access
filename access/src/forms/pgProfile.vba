Option Compare Database
Option Explicit

' Profile (see modBuildProfile.Build_pgProfile), a sidebar link for
' everyone signed in:
'   Your details  name, rank, sign-in name, email and command from the
'                 Personnel file; billet, unit, community, location, DSN
'                 and groups from the person's ANCHOR profile (tblPeople),
'                 which a System Admin keeps under Admin: Users.
'   Notices       Email notices (tblPeople.NotifyEmail), saved when flipped.
'   Roles         every ANCHOR role with what it allows, Granted or Not
'                 granted for this person.
'   Colors        the palette picker (AppAppearance).
'   Sign out      AppSignOut (modSignIn), which lets someone else sign in
'                 without closing Access.
' With no data file linked, or nobody signed in, the cards say so and the
' switch and Sign out are off.
'
' Owner: the Profile page's build task. Depends on: modErrors, modGlobals,
' modSignIn, modUI; the kit's modKitLocks, modKitUI, modKitUtil,
' modPersonnel.

Private Const DETAIL_ROWS As Long = 10          ' lblVal0..9 (modBuildProfile)

Private m_pid As Long           ' the signed-in person with a tblPeople row (0: none, or no data file)

Public Property Get HelpTopic() As String
    HelpTopic = "profile"
End Property

Private Sub Form_Load()
    Dim why As String
    On Error GoTo Err_Handler
    m_pid = 0
    ShowSignedIn
    If Not DataLinked() Then
        why = "This copy was built without the data file, so there is no profile to show."
    ElseIf SignedInPersonID() = 0 Then
        why = "Nobody is signed in, so there is no profile to show."
    ElseIf DCount("*", "tblPeople", "PersonID=" & SignedInPersonID()) = 0 Then
        why = "You aren't on ANCHOR's list of people yet. Ask a System Admin to add you under Admin: Users."
    Else
        m_pid = SignedInPersonID()
    End If
    If m_pid = 0 Then
        ShowNobody why
        Exit Sub
    End If
    ShowDetails
    Me.sfRoles.Form.RecordSource = "SELECT r.RoleID AS id, r.RoleName, r.Detail, IIf(pr.PersonID Is Null,'Not granted','Granted') AS granted " & _
        "FROM tblRoles AS r LEFT JOIN (SELECT PersonID, RoleID FROM tblPersonRoles WHERE PersonID=" & m_pid & ") AS pr " & _
        "ON r.RoleID=pr.RoleID ORDER BY r.SortOrder"
    ShowEmpty Me, "emptyRoles", Me.sfRoles.Form.RecordsetClone.RecordCount = 0
    SetVisible Me.sfRoles, Me.sfRoles.Form.RecordsetClone.RecordCount > 0
    SetSwitch "tgEmail", Nz(DLookup("NotifyEmail", "tblPeople", "PersonID=" & m_pid), False)
    ShowEmailLine
    Exit Sub
Err_Handler:
    HandleError "pgProfile.Form_Load"
End Sub

' The roles table (subGrid) calls this on a row click; nothing to open.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

' No profile to show: why, in the details card, and the switch locked.
Private Sub ShowNobody(ByVal why As String)
    Dim i As Long, c As Variant
    Me.lblName.Caption = "No profile"
    Me.lblBilletLine.Caption = Replace(why, "&", "&&")
    For i = 0 To DETAIL_ROWS - 1
        SetVisible Me.Controls("lblKey" & i), False
        SetVisible Me.Controls("lblVal" & i), False
    Next
    SetSwitch "tgEmail", False
    SetEnabled Me.tgEmail, False
    Me.tgEmail_val.Locked = True            ' the link beside it toggles the value too
    Me.lblEmailSub.Caption = "Email notices can be set once you're signed in."
    For Each c In Array("sfRoles", "lblRolesNote")
        SetVisible Me.Controls(c), False
    Next
    For i = 0 To 2
        SetVisible Me.Controls("hdrRoles" & i), False
    Next
    ShowEmpty Me, "emptyRoles", True
End Sub

' The details card, from the Personnel file and the ANCHOR profile.
Private Sub ShowDetails()
    Dim rs As DAO.Recordset, vals As Variant, i As Long, nm As String, subLine As String
    Set rs = CurrentDb.OpenRecordset("SELECT p.first_name, p.last_name, p.rank, p.username, p.email, p.command, e.Billet, e.Unit, " & _
        "e.Community, e.Location, e.Dsn, e.Groups, e.IsActive FROM tblPeople AS e INNER JOIN " & PERSON_TABLE & _
        " AS p ON e.PersonID=p.person_id WHERE e.PersonID=" & m_pid, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ShowNobody "Your record in the Personnel file can't be found. Ask a Personnel admin."
        m_pid = 0
        Exit Sub
    End If
    nm = Trim$(Nz(rs!rank, "") & " " & Nz(rs!first_name, "") & " " & Nz(rs!last_name, ""))
    Me.lblName.Caption = Replace(nm, "&", "&&")
    subLine = Nz(rs!Billet, "")
    If Len(Nz(rs!community, "")) > 0 Then subLine = subLine & IIf(Len(subLine) > 0, "  -  ", "") & rs!community
    If Len(subLine) = 0 Then subLine = "No billet recorded yet. A System Admin sets it under Admin: Users."
    Me.lblBilletLine.Caption = Replace(subLine, "&", "&&")
    vals = Array(rs!rank.Value, rs!username.Value, rs!email.Value, rs!command.Value, rs!Billet.Value, rs!Unit.Value, _
        rs!community.Value, rs!location.Value, rs!Dsn.Value, rs!Groups.Value)
    rs.Close
    For i = 0 To DETAIL_ROWS - 1
        If Len(Trim$(Nz(vals(i), ""))) = 0 Then
            Me.Controls("lblVal" & i).Caption = "-"
        Else
            Me.Controls("lblVal" & i).Caption = Replace(Trim$(vals(i)), "&", "&&")
        End If
    Next
End Sub

' The Sign out card: who is signed in, and its buttons only when there is
' someone signed in.
Private Sub ShowSignedIn()
    If DataLinked() And SignedInPersonID() <> 0 Then
        Me.lblSignedIn.Caption = "You're signed in as " & Replace(SignedInName(), "&", "&&") & _
            ". Sign out to let someone else sign in on this copy."
    Else
        Me.lblSignedIn.Caption = IIf(DataLinked(), "Nobody is signed in on this copy.", _
            "This copy was built without the data file, so nobody signs in.")
        SetEnabled Me.btnSignOut, False
        SetEnabled Me.btnChangePwd, False
    End If
End Sub

' The line under the switch: where emails go, or why none can.
Private Sub ShowEmailLine()
    Dim addr As String
    addr = Trim$(Nz(DLookup("email", PERSON_TABLE, "person_id=" & m_pid), ""))
    If Len(addr) > 0 Then
        Me.lblEmailSub.Caption = "An email to " & Replace(addr, "&", "&&") & " (the Personnel file's address) when traffic " & _
            "you must act on arrives."
        Me.lblEmailSub.ForeColor = Th(CLR_MUTED_FG)
    Else
        Me.lblEmailSub.Caption = "The Personnel file has no email address for you, so no email can reach you. Ask a Personnel admin."
        Me.lblEmailSub.ForeColor = Th(CLR_WARN_TEXT)
    End If
End Sub

' Email notices (Toggle's onChange): saved at once to the person's own row.
Public Sub EmailChanged()
    Dim v As Boolean
    On Error GoTo Err_Handler
    If m_pid = 0 Then Exit Sub
    v = IsSwitchOn("tgEmail")
    KitExecute "UPDATE tblPeople SET NotifyEmail=" & IIf(v, "True", "False") & " WHERE PersonID=" & m_pid
    KitToast "Email notices " & IIf(v, "on", "off"), "success"
    Exit Sub
Err_Handler:
    HandleError "pgProfile.EmailChanged", m_pid
    On Error Resume Next                    ' the save failed and was reported: show it as the data file has it
    SetSwitch "tgEmail", Nz(DLookup("NotifyEmail", "tblPeople", "PersonID=" & m_pid), False)
End Sub

Private Function IsSwitchOn(ByVal nm As String) As Boolean
    IsSwitchOn = CBool(Nz(Me.Controls(nm & "_val").Value, False))
End Function

' Set a switch in code (no onChange).
Private Sub SetSwitch(ByVal nm As String, ByVal v As Boolean)
    Me.Controls(nm & "_val").Value = v
    Me.Controls(nm & "_bg").Requery
End Sub
