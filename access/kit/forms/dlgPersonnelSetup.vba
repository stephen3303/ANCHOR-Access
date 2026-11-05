Option Compare Database
Option Explicit

' First-run Personnel setup (BuildPersonnelDialogs builds it; KitPersonnelStart
' opens it with RunDialog). Shows one panel at a time (controls tagged
' choice, create or ask) and fits the dialog to it. Ends "ok" when the
' Personnel file is linked: found, or created with the current Windows user
' signed in as its first admin. Closing any other way leaves the app to quit.
' After everyone was deleted from the directory (PERSONNEL_NO_ADMIN, an
' admin's purge) it offers creating the first admin in the linked file
' instead: Create asks only for the name, rank and password.

Private m_tops As Object          ' each tagged control's designed Top, by name
Private m_fullHeight As Long      ' the designed Detail height (the tallest panel)
Private m_noAdmin As Boolean      ' the linked file has no admin: Create makes the first one in it

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim lost As Boolean, c As Control
    Set m_tops = CreateObject("Scripting.Dictionary")
    For Each c In Me.Controls
        If Len(c.Tag) > 0 Then m_tops(c.Name) = c.Top
    Next
    m_fullHeight = Me.Section(acDetail).Height
    lost = (PersonnelUiSetupMode() = PERSONNEL_LOST)
    m_noAdmin = (PersonnelUiSetupMode() = PERSONNEL_NO_ADMIN)
    If m_noAdmin Then
        Me.lblTitle.Caption = "Create the first Personnel admin"
        Me.lblWhy.Caption = "Everyone was deleted from the Personnel directory, so it has no admin. If you are " & _
            "the admin, create your account as its first admin. Otherwise ask your admin."
        ' The file is linked and stays: Find leaves, and Create takes its place.
        Me.btnFind.Tag = "never": Me.lblFindHint.Tag = "never"
        m_tops("btnNotAdmin") = m_tops("btnCreate"): m_tops("lblNotAdminHint") = m_tops("lblCreateHint")
        m_tops("btnCreate") = m_tops("btnFind"): m_tops("lblCreateHint") = m_tops("lblFindHint")
        Me.btnCreate.Caption = "Create the first admin"
        Me.lblCreateHint.Caption = "You become the Personnel admin for every app that uses it."
        HideFileRows
    ElseIf lost Then
        Me.lblTitle.Caption = "The Personnel file can't be reached"
        Me.lblWhy.Caption = APP_NAME & " can't reach its Personnel file at " & PersonnelPath() & _
            ". If it has moved, find it. Otherwise ask your admin."
        ' Never offer a second directory while one is linked: Create leaves
        ' every panel, and I'm not the admin takes its place.
        Me.btnCreate.Tag = "never": Me.lblCreateHint.Tag = "never"
        m_tops("btnNotAdmin") = m_tops("btnCreate")
        m_tops("lblNotAdminHint") = m_tops("lblCreateHint")
    Else
        Me.lblWhy.Caption = APP_NAME & " signs people in from a shared Personnel file, and it can't find one. " & _
            "If Personnel is already set up for another app, find its file."
    End If
    Me.txtFolder = Left$(PersonnelDefaultPath(PersonnelUiAppTable()), InStrRev(PersonnelDefaultPath(PersonnelUiAppTable()), "\") - 1)
    Me.txtAsk = PersonnelAdminNote()
    FitWhy
    ShowPanel "choice"
    Exit Sub
Err_Handler:
    LogError "dlgPersonnelSetup.Form_Load"
End Sub

' The first admin goes into the linked file: no folder or site key to choose.
' Their rows leave the create panel, and the rows under them move up.
Private Sub HideFileRows()
    Dim v As Variant, shift As Long, c As Control
    shift = m_tops("lblPwd") - m_tops("lblFolder")
    For Each c In Me.Controls
        If c.Tag = "create" Then
            If m_tops(c.Name) >= m_tops("lblPwd") Then m_tops(c.Name) = m_tops(c.Name) - shift
        End If
    Next
    For Each v In Array("lblFolder", "txtFolder", "btnFolder", "lblFolderHint", "lblKey", "txtKey", "lblKeyHint")
        Me.Controls(v).Tag = "never"
    Next
End Sub

' F1 opens the help on this screen (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "directory.setup"
End Property

' F1 opens this dialog's help even during startup (PersonnelF1).
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    PersonnelF1 KeyCode
End Sub

' Show the controls tagged panel (and the untagged title) at their designed
' places, hide the rest, and fit the dialog to the panel. Access won't make
' the Detail shorter than its lowest control, hidden or not, so hidden
' controls move to the top first; the Detail grows back to its designed
' height before the shown ones return to their places.
Private Sub ShowPanel(ByVal panel As String)
    Dim c As Control
    Me.Section(acDetail).Height = m_fullHeight
    For Each c In Me.Controls
        If Len(c.Tag) > 0 Then
            If c.Tag <> panel Then
                SetVisible c, False
                c.Top = 0
            End If
        End If
    Next
    For Each c In Me.Controls
        If c.Tag = panel Then
            c.Top = m_tops(c.Name)
            SetVisible c, True
        End If
    Next
    FitToPanel panel
    Me.Repaint                                  ' no trace of the last panel's text
End Sub

' Make lblWhy as tall as its text (two lines usually, four or more with a
' long path) and move the choices under it to match. The lines are counted
' by wrapping the words as the label does, since a long path can't break at
' a space and leaves the line before it short.
Private Sub FitWhy()
    Dim lines As Long, dy As Long, k As Variant
    lines = WrapLines(Me.lblWhy.Caption, Me.lblWhy.Width - 60)    ' less the label's inner margins
    dy = lines * CLng(10 * 20 * 1.36) + 30 - Me.lblWhy.Height
    If dy = 0 Then Exit Sub
    m_fullHeight = m_fullHeight + IIf(dy > 0, dy, 0)
    Me.Section(acDetail).Height = m_fullHeight
    Me.lblWhy.Height = Me.lblWhy.Height + dy
    For Each k In m_tops.Keys
        If Me.Controls(k).Tag = "choice" And m_tops(k) > m_tops("lblWhy") Then m_tops(k) = m_tops(k) + dy
    Next
End Sub

' Lines text takes at 10pt in a label avail twips wide, wrapping word by
' word. A word wider than the line (a long path) breaks across lines.
Private Function WrapLines(ByVal text As String, ByVal avail As Long) As Long
    Dim para As Variant, wd As Variant, n As Long, used As Long, ww As Long, gap As Long
    gap = MeasureText("a b", 10, 400) - MeasureText("ab", 10, 400)
    For Each para In Split(Replace(text, vbCrLf, vbLf), vbLf)
        n = n + 1: used = 0
        For Each wd In Split(CStr(para), " ")
            If Len(wd) > 0 Then
                ww = MeasureText(CStr(wd), 10, 400)
                If used > 0 And used + gap + ww > avail Then n = n + 1: used = 0
                If used > 0 Then used = used + gap
                used = used + ww
                Do While used > avail
                    n = n + 1: used = used - avail
                Loop
            End If
        Next
    Next
    If n < 1 Then n = 1
    WrapLines = n
End Function

' Height to the lowest visible control of panel, plus the dialog padding.
Private Sub FitToPanel(ByVal panel As String)
    Dim c As Control, lowest As Long
    For Each c In Me.Controls
        If c.Tag = panel And c.Visible Then
            If c.Top + c.Height > lowest Then lowest = c.Top + c.Height
        End If
    Next
    Me.Section(acDetail).Height = lowest + Me.lblTitle.Top
    Me.InsideHeight = Me.Section(acDetail).Height
    CenterOnApp Me
End Sub

Private Sub btnFind_Click()
    On Error GoTo Err_Handler
    Dim path As String, problem As String
    Do
        path = KitPickFile("Find the Personnel file", "Personnel file|" & PERSONNEL_FILE & "|Access database|*.accdb", "personnel")
        If Len(path) = 0 Then Exit Sub
        problem = PersonnelFileProblem(path)
        If Len(problem) = 0 Then Exit Do
        KitMsgBox problem & vbCrLf & vbCrLf & "Pick the Personnel file again, or Cancel.", vbExclamation
    Loop
    LinkPersonnelFile path
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnFind_Click", path
End Sub

Private Sub btnCreate_Click()
    On Error GoTo Err_Handler
    ShowPanel "create"
    Me.lblTitle.Caption = IIf(m_noAdmin, "Create the first Personnel admin", "Create the Personnel file")
    If Len(Nz(Me.txtKey, "")) = 0 Then Me.txtKey = NewPersonnelKey()
    Me.txtFirst.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnCreate_Click"
End Sub

Private Sub btnNotAdmin_Click()
    On Error GoTo Err_Handler
    ShowPanel "ask"
    Me.lblTitle.Caption = "Ask your Personnel admin"
    Me.btnCopy.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnNotAdmin_Click"
End Sub

Private Sub btnBack_Click()
    On Error GoTo Err_Handler
    ShowPanel "choice"
    Me.lblTitle.Caption = IIf(m_noAdmin, "Create the first Personnel admin", "Set up the Personnel directory")
    Me.btnCreate.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnBack_Click"
End Sub

Private Sub btnFolder_Click()
    On Error GoTo Err_Handler
    Dim folder As String
    folder = KitPickFolder("Folder for the Personnel file", "personnel")
    If Len(folder) > 0 Then Me.txtFolder = folder
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnFolder_Click"
End Sub

' Create the file with the site key shown (or, with no admin, use the linked
' one), make the current Windows user its first admin with the password
' typed, and sign them in.
Private Sub btnCreateGo_Click()
    On Error GoTo Err_Handler
    Dim firstName As String, lastName As String, pwd As String, path As String, id As Long, problem As String, siteKey As String
    Me.btnCreateGo.SetFocus                     ' commits the box being typed in
    firstName = Trim$(Nz(Me.txtFirst, "")): lastName = Trim$(Nz(Me.txtLast, "")): pwd = Nz(Me.txtPwd, "")
    If Len(firstName) = 0 Or Len(lastName) = 0 Then KitMsgBox "Type your first and last name.", vbExclamation: Exit Sub
    If StrComp(pwd, Nz(Me.txtPwd2, ""), vbBinaryCompare) <> 0 Then KitMsgBox "The two passwords don't match.", vbExclamation: Exit Sub
    If Len(pwd) < 8 Then KitMsgBox "Use at least 8 characters for your password.", vbExclamation: Exit Sub
    If m_noAdmin Then
        DoCmd.Hourglass True
        id = AddFirstAdmin(PersonnelUiAppCode(), firstName, lastName, Trim$(Nz(Me.txtRank, "")))
    Else
        If Not NewFileChecked(path, siteKey) Then Exit Sub
        DoCmd.Hourglass True
        id = CreatePersonnelFile(path, PersonnelUiAppCode(), firstName, lastName, Trim$(Nz(Me.txtRank, "")), siteKey)
    End If
    problem = PersonnelSetPassword(id, pwd)
    DoCmd.Hourglass False
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation
    PersonnelSignInWindows PersonnelUiAppCode()
    PersonnelUiCreated                          ' KitPersonnelStart opens Add person next
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    DoCmd.Hourglass False
    HandleError "dlgPersonnelSetup.btnCreateGo_Click", path, IIf(m_noAdmin, "The first admin couldn't be created.", _
        "The Personnel file couldn't be created.")
End Sub

' The new file's path and site key from the create panel, checked: False,
' having told the user, when the key is unusable or the folder already has
' a Personnel file or key.
Private Function NewFileChecked(ByRef path As String, ByRef siteKey As String) As Boolean
    siteKey = Trim$(Nz(Me.txtKey, ""))
    If Len(PersonnelKeyProblem(siteKey)) > 0 Then KitMsgBox PersonnelKeyProblem(siteKey), vbExclamation: Exit Function
    path = Nz(Me.txtFolder, "")
    If Right$(path, 1) <> "\" Then path = path & "\"
    path = path & PERSONNEL_FILE
    If FileExists(path) Or FileExists(PersonnelKeyPath(path)) Then
        KitMsgBox "There is already a Personnel file or key in that folder. Go Back and choose Find the Personnel file.", vbExclamation
        Exit Function
    End If
    NewFileChecked = True
End Function

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If KitCopyText(Nz(Me.txtAsk, "")) Then KitToast "Copied", "success" Else KitToast "Couldn't copy. Try again.", "warning"
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnCopy_Click"
End Sub

Private Sub btnCloseApp_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnelSetup.btnCloseApp_Click"
End Sub
