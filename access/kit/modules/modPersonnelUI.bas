Option Compare Database
Option Explicit

' Access UI kit - the Personnel directory's screens (data and sign-in rules
' are in modPersonnel).
'
' Startup, straight after the app's own KitEnsureBackend:
'   If Not KitPersonnelStart("item", APP_CODE) Then DoCmd.Quit
' finds or sets up the Personnel file (dlgPersonnelSetup), then signs the
' Windows user in with no questions, or shows the sign-in screen (dlgSignIn)
' when their account isn't set up, is turned off, or has no access to the
' app. The sidebar's Sign out button:
'   If Not KitPersonnelSignOut(APP_CODE) Then DoCmd.Quit
' clears the session and shows the sign-in screen, so another person can
' sign in on the same copy. A Personnel link or button for admins:
'   =KitPersonnelManage("SABER")
' opens the Personnel page (dlgPersonnel): the list of people, Add person
' and Import (dlgPerson adds or changes one, and resets a password).
'
' To pick a person on a page or dialog (an assignee, a reviewer):
'   PersonPickBox "txtOwner", "owner_person_id", x, y, w
' a searchable picker (PickBox, modMenu) of the active people.
'
' Build the dialogs with the app's others: BuildPersonnelDialogs. Help
' topics for them are in kit\help\personnel.txt, to copy into the app's
' help file.

Public Const SETUP_FORM As String = "dlgPersonnelSetup"
Public Const SIGNIN_FORM As String = "dlgSignIn"
Public Const PERSONNEL_FORM As String = "dlgPersonnel"
Public Const PERSON_FORM As String = "dlgPerson"

' What the open Personnel dialog is for (the dialogs read these).
Private m_appTable As String
Private m_appCode As String
Private m_setupMode As Long
Private m_signInWhy As String
Private m_created As Boolean       ' the setup screen created the Personnel file

' ---------------------------------------------------------------------------
' Startup and sign out
' ---------------------------------------------------------------------------

' Make sure the Personnel file is linked and someone is signed in to appCode
' (see the top of the module). False when the user closed the app instead.
Public Function KitPersonnelStart(ByVal appTable As String, ByVal appCode As String) As Boolean
    Dim status As Long
    On Error GoTo Fail
    m_appTable = appTable: m_appCode = appCode: m_created = False
    status = PersonnelStatus(appTable)
    If status <> PERSONNEL_READY Then
        m_setupMode = status
        CloseSplash
        If Not RunDialog(SETUP_FORM) Then Exit Function
        If PersonnelStatus(appTable) <> PERSONNEL_READY Then Exit Function
    End If
    If SignedInPersonID() > 0 Then
        ' The admin who just created the file goes on to add the first people.
        ' Anyone else already signed in (after Find, too) just carries on.
        If m_created And PersonnelIsAdminNow() Then RunDialog PERSON_FORM, "new"
        KitPersonnelStart = True
        Exit Function
    End If
    SplashStatus "Checking your Windows sign-in" & ChrW(8230)
    If PersonnelSignInWindows(appCode) Then KitPersonnelStart = True: Exit Function
    m_signInWhy = PersonnelWindowsWhy(appCode)
    CloseSplash
    KitPersonnelStart = RunDialog(SIGNIN_FORM)
    Exit Function
Fail:
    HandleError "modPersonnelUI.KitPersonnelStart", appCode, "The Personnel directory couldn't be opened."
End Function

' Sign out and show the sign-in screen. False when the user closed the app
' instead of signing in.
Public Function KitPersonnelSignOut(ByVal appCode As String) As Boolean
    On Error GoTo Fail
    m_appCode = appCode
    PersonnelSignOut
    m_signInWhy = "signout"
    KitPersonnelSignOut = RunDialog(SIGNIN_FORM)
    Exit Function
Fail:
    HandleError "modPersonnelUI.KitPersonnelSignOut", appCode
End Function

' The Personnel page (dlgPersonnel), for directory admins; anyone else is
' told to ask one. For a sidebar link: =KitPersonnelManage("SABER").
Public Function KitPersonnelManage(ByVal appCode As String)
    On Error GoTo Fail
    If Not PersonnelIsAdminNow() Then                ' not the sign-in's copy: an admin may have changed it since
        KitMsgBox "Only a Personnel admin can add or change people. Ask one to make the change.", vbInformation, "Personnel"
        Exit Function
    End If
    m_appCode = appCode
    RunDialog PERSONNEL_FORM
    Exit Function
Fail:
    HandleError "modPersonnelUI.KitPersonnelManage", appCode, "The Personnel page couldn't be opened."
End Function

' Why the Windows account running Access can't sign in to appCode:
' "notset" (nobody's account), "off" (a person who is turned off),
' "noaccess" (an active person without access to appCode), or "" when it
' can. For an app's own sign-in form too, with PersonnelSignInWhyText.
Public Function PersonnelWindowsWhy(ByVal appCode As String) As String
    Dim id As Long
    id = PersonnelWindowsPersonId()
    If id = 0 Then
        PersonnelWindowsWhy = "notset"
    ElseIf DCount("*", PERSON_TABLE, "is_active AND person_id=" & id) = 0 Then
        PersonnelWindowsWhy = "off"
    ElseIf Not PersonCanUse(id, appCode) Then
        PersonnelWindowsWhy = "noaccess"
    End If
End Function

' The sign-in screen's line on why it shows, with who to ask
' (PersonnelAdminNames): for why (PersonnelWindowsWhy), or else for
' PersonnelUiSignInWhy. "" for "signout" or "".
Public Function PersonnelSignInWhyText(Optional ByVal why As Variant) As String
    If IsMissing(why) Then why = m_signInWhy
    Select Case why
        Case "notset"
            PersonnelSignInWhyText = "This Windows account isn't set up in Personnel yet. Ask a Personnel admin" & _
                PersonnelAdminNames() & " to add it, or sign in with your sign-in name and app password."
        Case "noaccess"
            PersonnelSignInWhyText = "This Windows account is in Personnel, but it doesn't have access to " & APP_NAME & _
                " yet. Ask a Personnel admin" & PersonnelAdminNames() & " to give you access."
        Case "off"
            PersonnelSignInWhyText = "This Windows account's Personnel record is turned off. Ask a Personnel admin" & _
                PersonnelAdminNames() & " to turn it back on."
    End Select
End Function

' For the dialogs: the app's code and linked table, why the setup screen is
' open (PERSONNEL_NOT_SET_UP or PERSONNEL_LOST), why the sign-in screen is
' ("notset", "noaccess" or "off": PersonnelWindowsWhy; "signout").
' A searchable picker (PickBox, modMenu) of the active people in the
' Personnel file: typing part of a last or first name, a rank or a sign-in
' name lists the matches as "Smith, John (LT)" with the sign-in name on the
' right. nm & "Key" holds the person_id, bound to keyField ("" for none).
' Build dlgPick with the app's dialogs (BuildPickDialog), and wire the box
' in the form's code as PickBox says.
Public Function PersonPickBox(ByVal nm As String, ByVal keyField As String, ByVal x As Long, ByVal y As Long, _
        ByVal w As Long, Optional ByVal h As Long = 450) As Control
    Set PersonPickBox = PickBox(nm, keyField, "SELECT person_id, last_name & ', ' & first_name & " & _
        "IIf(Nz(rank, '') = '', '', ' (' & rank & ')') AS person_name, username FROM " & PERSON_TABLE & _
        " WHERE is_active", x, y, w, h, "Type a name")
End Function

Public Function PersonnelUiAppCode() As String
    PersonnelUiAppCode = m_appCode
End Function

Public Function PersonnelUiAppTable() As String
    PersonnelUiAppTable = m_appTable
End Function

Public Function PersonnelUiSetupMode() As Long
    PersonnelUiSetupMode = m_setupMode
End Function

Public Function PersonnelUiSignInWhy() As String
    PersonnelUiSignInWhy = m_signInWhy
End Function

' F1 in a Personnel dialog (its Form_KeyDown): open the kit help on the
' dialog's topic. The dialogs can open from the startup form's Open, before
' Access has read the AutoKeys macro, when F1 would show Access's own help.
' KitHelp is called by name, so an app without modKitHelp still compiles.
Public Sub PersonnelF1(KeyCode As Integer)
    If KeyCode <> vbKeyF1 Then Exit Sub
    KeyCode = 0
    On Error Resume Next                        ' no modKitHelp: F1 does nothing
    Application.Run "KitHelp"
End Sub

' dlgPersonnelSetup: the Personnel file was created (not found), so
' KitPersonnelStart opens Add person next.
Public Sub PersonnelUiCreated()
    m_created = True
End Sub

' What a person who isn't the admin sends to their admin.
Public Function PersonnelAdminNote() As String
    Dim dataFile As String
    dataFile = KitBackendPath(m_appTable)
    If Len(dataFile) = 0 Then dataFile = CurrentProject.FullName
    PersonnelAdminNote = APP_NAME & " can't find its Personnel file." & vbCrLf & _
        "App data file: " & dataFile & vbCrLf & _
        "Personnel file looked for: " & IIf(Len(PersonnelPath()) > 0, PersonnelPath(), PersonnelDefaultPath(m_appTable)) & vbCrLf & _
        "Computer: " & WinComputer() & vbCrLf & _
        "Windows account: " & WinAccount()
End Function

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' Build step: dlgPersonnelSetup, dlgSignIn, dlgPersonnel and dlgPerson
' (code in kit\forms).
Public Sub BuildPersonnelDialogs()
    Build_dlgPersonnelSetup
    Build_dlgSignIn
    Build_dlgPersonnel
    Build_dlgPerson
End Sub

' dlgPersonnelSetup: three panels in one dialog, shown one at a time by Tag
' (the dialog sizes itself to the panel):
'   choice  why it shows and three choices, each with a one-line hint:
'           Find the Personnel file, Create a new Personnel file (never while
'           one is linked), I'm not the admin
'   create  the first admin's name and rank, the folder (Change opens the
'           folder picker), the site key (a new random one, filled in) and
'           their app password twice
'   ask     what to send the admin, with Copy, and Close app
Private Sub Build_dlgPersonnelSetup()
    Const DLG_W As Long = 7200          ' 480 px
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, half As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: half = (w - Sp(SP_ITEM_GAP)) \ 2: btnW = px(96)
    NewDialog "Personnel", DLG_W, px(480)
    Ev F, "OnLoad", "OnKeyDown": F.KeyPreview = True   ' F1 (PersonnelF1)
    y = pad
    Lbl "lblTitle", "Set up the Personnel directory", x, y, w, px(26), 13, True
    y = y + px(26) + gap
    Set c = Lbl("lblWhy", "", x, y, w, px(72))            ' filled in and fitted to its text by Form_Load
    c.Tag = "choice"
    ' choice
    y = y + px(72) + gap
    ChoiceRow "Find", "Find the Personnel file", _
        "When Personnel is already set up and you know where its file is.", x, y, w
    ChoiceRow "Create", "Create a new Personnel file", _
        "For the first app set up here. You become the Personnel admin.", x, y, w
    ChoiceRow "NotAdmin", "I'm not the admin", _
        "Shows what to send your admin, then closes " & APP_NAME & ".", x, y, w
    ' create: starts under the title
    y = pad + px(26) + gap
    FieldLabel "lblFirst", "First name", x, y, half, "create"
    FieldLabel "lblLast", "Last name", x + half + Sp(SP_ITEM_GAP), y, half, "create"
    y = y + px(18)
    Txt("txtFirst", "", x, y, half, px(30)).Tag = "create"
    Txt("txtLast", "", x + half + Sp(SP_ITEM_GAP), y, half, px(30)).Tag = "create"
    y = y + px(30) + gap
    FieldLabel "lblRank", "Rank (optional)", x, y, half, "create"
    y = y + px(18)
    Txt("txtRank", "", x, y, half, px(30)).Tag = "create"
    y = y + px(30) + gap
    FieldLabel "lblFolder", "Folder for the Personnel file", x, y, w, "create"
    y = y + px(18)
    Set c = RO("txtFolder", "", x, y, w - btnW - Sp(SP_ITEM_GAP), px(30), 10)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE: c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8)
    c.Tag = "create"
    Btn("btnFolder", "Change...", x + w - btnW, y, btnW, px(30), "outline").Tag = "create"
    y = y + px(30) + px(4)
    Lbl("lblFolderHint", "Only Personnel admins should be able to change this folder. Ask IT to restrict it.", _
        x, y, w, px(18), 9, False, CLR_MUTED_FG).Tag = "create"
    y = y + px(18) + gap
    FieldLabel "lblKey", "Site key", x, y, w, "create"
    y = y + px(18)
    Txt("txtKey", "", x, y, w, px(30)).Tag = "create"                   ' btnCreate_Click fills in a new random key
    y = y + px(30) + px(4)
    Lbl("lblKeyHint", "Saved as " & PERSONNEL_KEY_FILE & " beside the file for every app. " & _
        "Keep it unless your site has one.", _
        x, y, w, px(32), 9, False, CLR_MUTED_FG).Tag = "create"
    y = y + px(32) + gap
    FieldLabel "lblPwd", "Your app password", x, y, half, "create"
    FieldLabel "lblPwd2", "Type it again", x + half + Sp(SP_ITEM_GAP), y, half, "create"
    y = y + px(18)
    Set c = Txt("txtPwd", "", x, y, half, px(30)): c.InputMask = SECRET_MASK: c.Tag = "create"
    Set c = Txt("txtPwd2", "", x + half + Sp(SP_ITEM_GAP), y, half, px(30)): c.InputMask = SECRET_MASK: c.Tag = "create"
    y = y + px(30) + px(4)
    Lbl("lblPwdHint", "8 or more characters. Capitals count.", x, y, w, px(18), 9, False, CLR_MUTED_FG).Tag = "create"
    y = y + px(18) + Sp(SP_DIALOG)
    Set c = Btn("btnCreateGo", "Create", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32)): c.Tag = "create"
    Set c = Btn("btnBack", "Back", x + w - btnW, y, btnW, px(32), "outline"): c.Tag = "create"
    ' ask: starts under the title
    y = pad + px(26) + gap
    Lbl("lblAsk", "Send this to your Personnel admin. They can find the file, or set Personnel up.", _
        x, y, w, px(36)).Tag = "ask"
    y = y + px(36) + gap
    Set c = RO("txtAsk", "", x, y, w, px(140), 9)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE: c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8): c.TopMargin = px(6): c.RightMargin = px(8)
    c.Tag = "ask"
    y = y + px(140) + Sp(SP_DIALOG)
    Set c = Btn("btnCopy", "Copy", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32), "outline"): c.Tag = "ask"
    Set c = Btn("btnCloseApp", "Close app", x + w - btnW, y, btnW, px(32)): c.Tag = "ask"
    ' Tall enough for the tallest panel; Form_Load fits it to the one shown.
    For Each c In F.Controls
        If c.Top + c.Height > y Then y = c.Top + c.Height
    Next
    F.Section(acDetail).Height = y + pad
    SaveForm SETUP_FORM
End Sub

' One choice: an outline button the width of the dialog and a muted hint
' under it; moves y below them.
Private Sub ChoiceRow(ByVal key As String, ByVal cap As String, ByVal hint As String, ByVal x As Long, y As Long, ByVal w As Long)
    Dim c As Control
    Set c = Btn("btn" & key, cap, x, y, w, px(36), "outline")
    c.Tag = "choice"
    y = y + px(36) + px(4)
    Lbl("lbl" & key & "Hint", hint, x, y, w, px(18), 9, False, CLR_MUTED_FG).Tag = "choice"
    y = y + px(18) + Sp(SP_ROW_GAP)
End Sub

Private Sub FieldLabel(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal tg As String)
    Lbl(nm, cap, x, y, w, px(16), 9, True, CLR_MUTED_FG).Tag = tg
End Sub

' dlgSignIn: why it shows (this Windows account isn't set up, is turned off
' or has no access to the app, with who to ask, the account and Copy; or
' signed out), sign-in name and app password, "Use my Windows sign-in next
' time" (only while this Windows account isn't saved for anyone), a line for
' what went wrong, and Sign in / Close app. Rows that don't apply are hidden
' and the rows under them move up (Form_Load).
Private Sub Build_dlgSignIn()
    Const DLG_W As Long = 6000          ' 400 px
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: btnW = px(96)
    NewDialog "Sign in", DLG_W, px(400)
    F.CloseButton = False                           ' Close app or Sign in ends it
    Ev F, "OnLoad", "OnKeyDown": F.KeyPreview = True   ' F1 (PersonnelF1)
    y = pad
    Lbl "lblTitle", "Sign in to " & APP_NAME, x, y, w, px(26), 13, True
    y = y + px(26) + gap
    ' Row "why": why this Windows account wasn't signed in. Form_Load sets the
    ' text (PersonnelSignInWhyText) and fits the height to it.
    Lbl("lblWhy", "This Windows account isn't set up in Personnel yet. Ask a Personnel admin to add it, " & _
        "or sign in with your sign-in name and app password.", x, y, w, px(52), 9).Tag = "why"
    y = y + px(52) + px(4)
    Set c = RO("txtAccount", "", x, y, w - px(72) - Sp(SP_ITEM_GAP), px(28), 9)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE: c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8): c.Tag = "why"
    Btn("btnCopyAccount", "Copy", x + w - px(72), y, px(72), px(28), "outline").Tag = "why"
    y = y + px(28) + gap
    Lbl "lblUser", "Sign-in name", x, y, w, px(16), 9, True, CLR_MUTED_FG
    y = y + px(18)
    Txt "txtUser", "", x, y, w, px(30)
    y = y + px(30) + gap
    Lbl "lblPwd", "App password", x, y, w, px(16), 9, True, CLR_MUTED_FG
    y = y + px(18)
    Set c = Txt("txtPwd", "", x, y, w, px(30)): c.InputMask = SECRET_MASK
    y = y + px(30) + gap
    ' Row "win": save this Windows account to the person signing in.
    Chk("chkWindows", "", x, y + px(2)).Tag = "win"
    Lbl("lblWindows", "Use my Windows sign-in next time", x + px(24), y, w - px(24), px(20), 9).Tag = "win"
    y = y + px(20) + gap
    Set c = Lbl("lblError", "", x, y, w, px(36), 9, False, CLR_ERROR)
    y = y + px(36) + gap
    Set c = Btn("btnSignIn", "Sign in", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32)): c.Default = True
    Set c = Btn("btnClose", "Close app", x + w - btnW, y, btnW, px(32), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    SaveForm SIGNIN_FORM
End Sub

' dlgPersonnel, the Personnel page for directory admins: a search box, Import
' and Add person, the list of people (double-click or Edit to change one)
' and Close.
Private Sub Build_dlgPersonnel()
    Const DLG_W As Long = 11520         ' 768 px
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, btnW As Long, listH As Long, colW As String
    pad = Sp(SP_DIALOG)
    x = pad: w = DLG_W - 2 * pad: btnW = px(110): listH = px(340)
    NewDialog "Personnel", DLG_W, px(540)
    Ev F, "OnLoad", "OnOpen", "OnKeyDown": F.KeyPreview = True   ' Form_Open checks admin and themes the form; F1 (PersonnelF1)
    y = pad
    Lbl "lblTitle", "Personnel", x, y, w - px(260), px(26), 13, True
    y = y + px(26) + Sp(16)
    Set c = Txt("txtSearch", "", x, y, px(260), px(30), 9)
    c.Format = "@;""Search people..."""
    Ev c, "OnChange"
    Btn "btnImport", "Import...", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(30), "outline"
    Btn "btnAdd", "Add person", x + w - btnW, y, btnW, px(30)
    y = y + px(30) + Sp(SP_ROW_GAP)
    ' person_id, Name, Sign-in name, Windows account, Can open, Status
    colW = "0;" & px(230) & ";" & px(120) & ";" & px(170) & ";" & px(80) & ";" & px(80)
    Set c = Lst("lstPeople", "", x, y, w, listH, 6, colW)
    Ev c, "OnDblClick"
    Set c = Lbl("lblNone", "No people match your search.", x + px(10), y + px(30), px(300), px(20), 9, False, CLR_MUTED_FG)
    c.Visible = False
    y = y + listH + Sp(16)
    Lbl "lblTip", "Double-click a person to change their details, reset their password or take away access.", _
        x, y, w - 2 * px(96) - Sp(SP_ITEM_GAP) - Sp(16), px(30), 8, False, CLR_MUTED_FG
    Btn "btnEdit", "Edit", x + w - 2 * px(96) - Sp(SP_ITEM_GAP), y, px(96), px(30), "outline"
    Set c = Btn("btnClose", "Close", x + w - px(96), y, px(96), px(30), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(30) + pad
    SaveForm PERSONNEL_FORM
End Sub

' dlgPerson: add or change one person. OpenArgs "new" or a person_id. Name,
' rank, command and email; sign-in name and Windows account; Active,
' Personnel admin and "Can open <app>"; an app password typed twice; a line
' for what went wrong; Save and Close. Adding stays open after Save, ready
' for the next person.
Private Sub Build_dlgPerson()
    Const DLG_W As Long = 8400          ' 560 px
    Dim c As Control, pad As Long, w As Long, x As Long, x2 As Long, y As Long, half As Long, gap As Long, btnW As Long
    pad = Sp(SP_DIALOG): gap = Sp(SP_ROW_GAP)
    x = pad: w = DLG_W - 2 * pad: half = (w - Sp(16)) \ 2: x2 = x + half + Sp(16): btnW = px(96)
    NewDialog "Person", DLG_W, px(640)
    Ev F, "OnLoad", "OnOpen", "OnKeyDown": F.KeyPreview = True   ' Form_Open checks admin and themes the form; F1 (PersonnelF1)
    y = pad
    Lbl "lblTitle", "Add person", x, y, w, px(26), 13, True
    y = y + px(26) + Sp(16)
    PersonRow "First", "First name", "Last", "Last name", x, x2, y, half
    PersonRow "Middle", "Middle initial", "Rank", "Rank", x, x2, y, half
    PersonRow "Command", "Command", "Email", "Email", x, x2, y, half
    PersonRow "User", "Sign-in name", "Account", "Windows account (DOMAIN\user)", x, x2, y, half
    y = y - gap + px(4)                             ' the hint sits under the sign-in name row
    Lbl "lblUserHint", "Blank: made from the Windows account, or the first initial and last name.", _
        x, y, w, px(16), 8, False, CLR_MUTED_FG
    y = y + px(16) + Sp(16)
    Chk "chkActive", "", x, y + px(2)
    Lbl "lblActive", "Active (can sign in)", x + px(24), y, half - px(24), px(20), 9
    Chk "chkAdmin", "", x2, y + px(2)
    Lbl "lblAdmin", "Personnel admin", x2 + px(24), y, half - px(24), px(20), 9
    y = y + px(20) + gap
    Chk "chkApp", "", x, y + px(2)
    Lbl "lblApp", "Can open " & APP_NAME, x + px(24), y, w - px(24), px(20), 9
    y = y + px(20) + Sp(16)
    FieldLabel "lblPwd", "App password", x, y, half, ""
    FieldLabel "lblPwd2", "Type it again", x2, y, half, ""
    y = y + px(18)
    Set c = Txt("txtPwd", "", x, y, half, px(30)): c.InputMask = SECRET_MASK
    Set c = Txt("txtPwd2", "", x2, y, half, px(30)): c.InputMask = SECRET_MASK
    y = y + px(30) + px(4)
    Lbl "lblPwdHint", "", x, y, w, px(32), 8, False, CLR_MUTED_FG         ' Form_Load: new or change
    y = y + px(32) + gap
    Lbl "lblError", "", x, y, w, px(36), 9, False, CLR_ERROR
    y = y + px(36) + gap
    Set c = Btn("btnSave", "Save", x + w - 2 * btnW - Sp(SP_ITEM_GAP), y, btnW, px(32)): c.Default = True
    Set c = Btn("btnClose", "Close", x + w - btnW, y, btnW, px(32), "outline"): c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    SaveForm PERSON_FORM
End Sub

' Two labelled boxes side by side (txt<k1>, txt<k2>); moves y below them.
Private Sub PersonRow(ByVal k1 As String, ByVal cap1 As String, ByVal k2 As String, ByVal cap2 As String, _
        ByVal x As Long, ByVal x2 As Long, y As Long, ByVal half As Long)
    FieldLabel "lbl" & k1, cap1, x, y, half, ""
    FieldLabel "lbl" & k2, cap2, x2, y, half, ""
    y = y + px(18)
    Txt "txt" & k1, "", x, y, half, px(30)
    Txt "txt" & k2, "", x2, y, half, px(30)
    y = y + px(30) + Sp(SP_ROW_GAP)
End Sub
