Option Compare Database
Option Explicit

' Sign-in screen (BuildPersonnelDialogs builds it; KitPersonnelStart and
' KitPersonnelSignOut open it with RunDialog). Checks a sign-in name and app
' password (PersonnelSignIn). A person with no password yet chooses one here
' when this Windows account is theirs; anyone else is told to ask an admin.
' Rows tagged why (why this Windows account wasn't signed in, and who to
' ask) and win (save it for next time) show only when they apply; the rows
' under a hidden one move up, and under the why text as it grows or shrinks.
' Ends "ok" when someone is signed in; Close app leaves the app to quit.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.txtAccount = WinAccount()
    If Len(PersonnelSignInWhyText()) = 0 Then CollapseRow "why" Else FitWhy PersonnelSignInWhyText()
    If WindowsAccountKnown() Or Len(WinAccount()) = 0 Then CollapseRow "win"
    CenterOnApp Me
    Me.txtUser.SetFocus
    Exit Sub
Err_Handler:
    LogError "dlgSignIn.Form_Load"
End Sub

' F1 opens the help on this screen (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "directory.signin"
End Property

' F1 opens this dialog's help even during startup (PersonnelF1).
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    PersonnelF1 KeyCode
End Sub

' Hide the controls tagged row and move everything under them up by the
' space they took, then shrink the dialog by the same amount.
Private Sub CollapseRow(ByVal row As String)
    Dim c As Control, rowTop As Long, rowBottom As Long, dy As Long
    rowTop = -1
    For Each c In Me.Controls
        If c.Tag = row Then
            If rowTop < 0 Or c.Top < rowTop Then rowTop = c.Top
            If c.Top + c.Height > rowBottom Then rowBottom = c.Top + c.Height
        End If
    Next
    If rowTop < 0 Then Exit Sub
    dy = rowBottom - rowTop + Sp(SP_ROW_GAP)
    For Each c In Me.Controls
        If c.Tag = row Then
            SetVisible c, False
        ElseIf c.Top >= rowBottom Then
            c.Top = c.Top - dy
        End If
    Next
    Me.Section(acDetail).Height = Me.Section(acDetail).Height - dy
    Me.InsideHeight = Me.Section(acDetail).Height
End Sub

' Show text in lblWhy, as tall as its lines (9pt, about 17px a line), and
' move everything under it by the difference.
Private Sub FitWhy(ByVal text As String)
    Dim c As Control, oldBottom As Long, lines As Long, dy As Long
    Me.lblWhy.Caption = text
    lines = -Int(-(MeasureText(text, 9, 400) * 1.1) / Me.lblWhy.Width)
    If lines < 1 Then lines = 1
    dy = lines * px(17) + px(1) - Me.lblWhy.Height
    If dy = 0 Then Exit Sub
    oldBottom = Me.lblWhy.Top + Me.lblWhy.Height
    If dy > 0 Then Me.Section(acDetail).Height = Me.Section(acDetail).Height + dy   ' room first
    For Each c In Me.Controls
        If c.Top >= oldBottom Then c.Top = c.Top + dy
    Next
    Me.lblWhy.Height = Me.lblWhy.Height + dy
    If dy < 0 Then Me.Section(acDetail).Height = Me.Section(acDetail).Height + dy
    Me.InsideHeight = Me.Section(acDetail).Height
End Sub

Private Sub btnSignIn_Click()
    On Error GoTo Err_Handler
    Dim user As String, pwd As String, problem As String
    Me.btnSignIn.SetFocus                       ' commits the box being typed in
    user = Trim$(Nz(Me.txtUser, "")): pwd = Nz(Me.txtPwd, "")
    Me.lblError.Caption = ""
    If Len(user) = 0 Then Me.lblError.Caption = "Type your sign-in name.": Me.txtUser.SetFocus: Exit Sub
    If PersonnelNeedsPassword(user) Then
        If Not ChooseFirstPassword(user) Then Exit Sub
        pwd = Nz(Me.txtPwd, "")
    End If
    problem = PersonnelSignIn(user, pwd, PersonnelUiAppCode())
    If Len(problem) > 0 Then
        Me.lblError.Caption = problem
        Me.txtPwd = Null
        Me.txtPwd.SetFocus
        Exit Sub
    End If
    If Me.chkWindows.Visible And Nz(Me.chkWindows, False) Then PersonnelSaveWindowsAccount
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgSignIn.btnSignIn_Click", user
End Sub

' A person with no app password yet: let them choose one when this Windows
' account is theirs (asked twice, masked), and put it in the password box.
' False, with the reason shown, when they can't or cancelled.
Private Function ChooseFirstPassword(ByVal user As String) As Boolean
    Dim p1 As Variant, p2 As Variant, problem As String
    If Not PersonnelCanSetFirstPassword(user) Then
        Me.lblError.Caption = "You don't have an app password yet. Ask a Personnel admin to set one."
        Exit Function
    End If
    p1 = KitInputBox("You don't have an app password yet. Choose one (8 or more characters).", "Choose a password", , True)
    If IsNull(p1) Then Exit Function
    p2 = KitInputBox("Type the same password again.", "Choose a password", , True)
    If IsNull(p2) Then Exit Function
    If StrComp(p1, p2, vbBinaryCompare) <> 0 Then Me.lblError.Caption = "The two passwords don't match. Try again.": Exit Function
    problem = PersonnelSetPassword(PersonIdFor(user), p1)
    If Len(problem) > 0 Then Me.lblError.Caption = problem: Exit Function
    Me.txtPwd = p1
    ChooseFirstPassword = True
End Function

Private Sub btnCopyAccount_Click()
    On Error GoTo Err_Handler
    If KitCopyText(WinAccount()) Then KitToast "Copied", "success" Else KitToast "Couldn't copy. Try again.", "warning"
    Exit Sub
Err_Handler:
    HandleError "dlgSignIn.btnCopyAccount_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgSignIn.btnClose_Click"
End Sub
