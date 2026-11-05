Option Compare Database
Option Explicit

' Add or change one person in the Personnel file (BuildPersonnelDialogs
' builds it; dlgPersonnel and KitPersonnelStart open it). OpenArgs: "new" or
' a person_id. Adding stays open after Save, cleared for the next person;
' changing closes. Reset password gives someone else a one-time password.
' The UPN box shows and saves their user principal name (kit_person_upn,
' admins only); with no Windows account, Save looks it up from the UPN.
' Ends "ok" when anything was saved.

Private m_id As Long            ' 0 while adding
Private m_saved As Boolean

' Only an active Personnel admin may open this, whoever opens it: the
' check is against the Personnel file now (PersonnelIsAdminNow). RunDialog
' treats the cancelled open as not ok. The builder wires OnOpen to this, in
' place of the usual =ThemeForm([Form]) hook, so it themes the form too.
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler
    ThemeForm Me
    If Not PersonnelIsAdminNow() Then
        KitMsgBox "Only an active Personnel admin can add or change people.", vbExclamation
        Cancel = True
    End If
    Exit Sub
Err_Handler:
    Cancel = True
    HandleError "dlgPerson.Form_Open"
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim rs As DAO.Recordset
    If Nz(Me.OpenArgs, "new") = "new" Then
        StartNew
    Else
        m_id = CLng(Me.OpenArgs)
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & PERSON_TABLE & " WHERE person_id=" & m_id, dbOpenSnapshot)
        Me.lblTitle.Caption = "Change " & Nz(rs!display_name, rs!username)
        Me.txtFirst = rs!first_name: Me.txtLast = rs!last_name
        Me.txtMiddle = rs!middle_initial: Me.txtRank = rs!rank
        Me.txtCommand = rs!command: Me.txtEmail = rs!email
        Me.txtUser = rs!username: Me.txtAccount = rs!windows_account
        Me.txtUpn = PersonnelUpn(m_id)
        Me.chkActive = CBool(rs!is_active): Me.chkAdmin = CBool(rs!is_admin)
        If m_id = SignedInPersonID() Then
            Me.lblPwdHint.Caption = "Type a new password to change yours, or leave it blank to keep it."
            SetVisible Me.btnResetPwd, False        ' an admin's own: Change my password
        Else
            Me.lblPwdHint.Caption = IIf(IsNull(rs!password_hash), "No app password yet: they sign in with Windows. ", "") & _
                "Reset password makes a one-time password, or type one. They choose their own when they next sign in with it."
        End If
        rs.Close
        Me.chkApp = CBool(Nz(DLookup("is_active", PERSON_APP_TABLE, "person_id=" & m_id & " AND app_code='" & _
            Replace(PersonnelUiAppCode(), "'", "''") & "'"), False))
        SetVisible Me.lblUserHint, False
        CloseUpUserHint
    End If
    Me.lblApp.Caption = "Can open " & APP_NAME
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgPerson.Form_Load"
End Sub

' Changing someone has no sign-in hint: move everything under it up, so the
' UPN row sits one row gap under the sign-in name like every other row.
Private Sub CloseUpUserHint()
    Dim c As Control, hintTop As Long, delta As Long
    hintTop = Me.lblUserHint.Top
    delta = Me.lblUpn.Top - (Me.txtUser.Top + Me.txtUser.Height) - Sp(SP_ROW_GAP)
    If delta <= 0 Then Exit Sub
    For Each c In Me.Controls
        If c.Top > hintTop Then c.Top = c.Top - delta
    Next
    Me.Section(acDetail).Height = Me.Section(acDetail).Height - delta
    Me.InsideHeight = Me.InsideHeight - delta
End Sub

' F1 opens the help on adding and changing people (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "directory.person"
End Property

' F1 opens this dialog's help even during startup (PersonnelF1).
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    PersonnelF1 KeyCode
End Sub

' Empty boxes for the next person to add.
Private Sub StartNew()
    Dim nm As Variant
    m_id = 0
    Me.lblTitle.Caption = "Add person"
    For Each nm In Array("txtFirst", "txtLast", "txtMiddle", "txtRank", "txtCommand", "txtEmail", "txtUser", "txtAccount", "txtUpn", "txtPwd", "txtPwd2")
        Me.Controls(nm).Value = Null
    Next
    Me.chkActive = True: Me.chkAdmin = False: Me.chkApp = True
    Me.lblPwdHint.Caption = "Leave it blank when they sign in with Windows. Otherwise set one (8 or more characters) and tell them; " & _
        "they choose their own when they first sign in with it."
    SetVisible Me.btnResetPwd, False                ' nobody to reset yet
    Me.lblError.Caption = ""
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    Dim pwd As String, problem As String, id As Long, who As String, linked As Boolean
    Me.btnSave.SetFocus                         ' commits the box being typed in
    Me.lblError.Caption = ""
    PersonnelRequireAdmin                       ' error 5, shown below, if an admin has since changed them
    If m_id <> 0 And m_id = SignedInPersonID() And Not Nz(Me.chkApp, False) Then
        ' Checked before anything is saved; SetPersonApp refuses it too.
        Me.lblError.Caption = "You can't take away your own access to " & APP_NAME & ".": Exit Sub
    End If
    pwd = Nz(Me.txtPwd, "")
    If StrComp(pwd, Nz(Me.txtPwd2, ""), vbBinaryCompare) <> 0 Then   ' case matters; <> ignores it here
        Me.lblError.Caption = "The two passwords don't match."
        Exit Sub
    End If
    If Len(pwd) > 0 And Len(pwd) < 8 Then Me.lblError.Caption = "Use at least 8 characters for the password.": Exit Sub
    problem = PersonnelUpnProblem(m_id, Nz(Me.txtUpn, ""))
    If Len(problem) > 0 Then Me.lblError.Caption = problem: Exit Sub
    If m_id = 0 Then
        id = AddPerson(Nz(Me.txtFirst, ""), Nz(Me.txtLast, ""), Nz(Me.txtRank, ""), Nz(Me.txtAccount, ""), _
            Nz(Me.txtUser, ""), Nz(Me.chkAdmin, False), Nz(Me.txtMiddle, ""), Nz(Me.txtCommand, ""), Nz(Me.txtEmail, ""))
        If Not Nz(Me.chkActive, True) Then
            UpdatePerson id, Nz(Me.txtFirst, ""), Nz(Me.txtLast, ""), Nz(Me.txtRank, ""), Nz(Me.txtAccount, ""), _
                DLookup("username", PERSON_TABLE, "person_id=" & id), Nz(Me.chkAdmin, False), False, _
                Nz(Me.txtMiddle, ""), Nz(Me.txtCommand, ""), Nz(Me.txtEmail, "")
        End If
    Else
        id = m_id
        UpdatePerson id, Nz(Me.txtFirst, ""), Nz(Me.txtLast, ""), Nz(Me.txtRank, ""), Nz(Me.txtAccount, ""), _
            Nz(Me.txtUser, ""), Nz(Me.chkAdmin, False), Nz(Me.chkActive, False), _
            Nz(Me.txtMiddle, ""), Nz(Me.txtCommand, ""), Nz(Me.txtEmail, "")
    End If
    SetPersonApp id, PersonnelUiAppCode(), Nz(Me.chkApp, False)
    problem = PersonnelSetUpn(id, Nz(Me.txtUpn, ""))
    If Len(problem) = 0 And Len(Trim$(Nz(Me.txtAccount, ""))) = 0 Then linked = PersonnelLinkFromUpn(id)
    ' A password typed for someone else is known to the admin, so it is one-time.
    If Len(pwd) > 0 And Len(problem) = 0 Then problem = PersonnelSetPassword(id, pwd, id <> SignedInPersonID())
    m_saved = True
    TempVars("DialogResult") = "ok"
    who = Nz(DLookup("display_name", PERSON_TABLE, "person_id=" & id), "")
    If Len(problem) > 0 Then
        m_id = id                               ' saved: a second Save changes them, not adds them again
        Me.lblError.Caption = problem
        Exit Sub
    End If
    If linked Then who = who & ", linked to " & DLookup("windows_account", PERSON_TABLE, "person_id=" & id)
    If m_id = 0 Then
        KitToast "Added " & who, "success"
        StartNew
        Me.txtFirst.SetFocus
    Else
        KitToast "Saved " & who, "success"
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    Exit Sub
Err_Handler:
    If Err.Number = 5 Then Me.lblError.Caption = Err.Description: Exit Sub   ' a refusal: say why, keep what was typed
    HandleError "dlgPerson.btnSave_Click", CStr(m_id), "The person couldn't be saved."
End Sub

' A new one-time password for this person, shown to the admin once and put
' on the clipboard. The person must choose their own when they sign in with it.
Private Sub btnResetPwd_Click()
    On Error GoTo Err_Handler
    Dim who As String, pwd As String
    If m_id = 0 Then Exit Sub
    who = Nz(DLookup("display_name", PERSON_TABLE, "person_id=" & m_id), "this person")
    If KitMsgBox("Make a new one-time password for " & who & "? Their current app password stops working.", _
            vbQuestion + vbYesNo, "Reset password") <> vbYes Then Exit Sub
    pwd = PersonnelResetPassword(m_id)
    m_saved = True
    TempVars("DialogResult") = "ok"
    Me.txtPwd = Null: Me.txtPwd2 = Null
    KitCopyText pwd
    KitMsgBox "One-time password for " & who & ":" & vbCrLf & vbCrLf & pwd & vbCrLf & vbCrLf & _
        "It is on the clipboard until you close this message. Give it to them in person. When they sign in " & _
        "with it they choose their own. It isn't shown again.", vbInformation, "Reset password"
    KitCopyText " "                                 ' don't leave it on the clipboard
    Exit Sub
Err_Handler:
    If Err.Number = 5 Then Me.lblError.Caption = Err.Description: Exit Sub
    HandleError "dlgPerson.btnResetPwd_Click", CStr(m_id), "The password couldn't be reset."
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    If m_saved Then TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgPerson.btnClose_Click"
End Sub
