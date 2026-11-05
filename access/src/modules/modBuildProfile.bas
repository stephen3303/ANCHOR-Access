Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Profile page (/profile): the signed-in
' person's details and roles, email notices, colors and Sign out. Its
' code-behind is forms\pgProfile.vba.
'
' On the left the person's details (from the Personnel file and their ANCHOR
' profile) and the email notices switch; on the right their roles, granted
' or not, then Colors and Sign out, as Exam Generator's My settings.
'
' Owner: the Profile page's build task. Depends on: modBuild (PAGE_W,
' PAGE_H, PAGE_PAD, ContentW); the kit's modUIKit.

Private Const LEFT_W As Long = 7200             ' the left column, 480 px
Private Const DETAIL_ROWS As Long = 10          ' the details card's label / value rows

Private Function RightW() As Long
    RightW = ContentW() - LEFT_W - px(12)
End Function

' modBuild.BuildPages: the table first, then the page.
Public Sub BuildProfilePage()
    BuildGrid "subGridProfileRoles", RoleCols(), RightW() - px(32)
    Build_pgProfile
End Sub

' "field|caption|width px|options[|sort field]"
Private Function RoleCols() As Variant
    RoleCols = Array("RoleName|Role|156|b", "Detail|What it allows|212|m", "granted|Access|100|p")
End Function

Public Sub Build_pgProfile()
    Dim c As Control, i As Long, x As Long, y As Long, w As Long, fw As Long, h As Long, keys As Variant
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Profile", "Your details and roles, your email notices, the colors the app uses on this computer, " & _
        "and signing out. Your roles are changed by a System Admin under Admin: Users.", PAGE_PAD

    ' Your details.
    x = PAGE_PAD: y = px(84): w = LEFT_W: fw = w - px(32): h = px(332)
    Card "Details", "Your details", x, y, w, h
    Lbl "lblName", "", x + px(16), y + px(40), fw, px(26), 13, True
    Lbl "lblBilletLine", "", x + px(16), y + px(66), fw, px(20), 9, False, CLR_MUTED_FG
    keys = Array("Rank", "Signs in as", "Email", "Command", "Billet", "Unit", "Community", "Location", "Phone (DSN)", "Groups")
    For i = 0 To DETAIL_ROWS - 1
        Lbl "lblKey" & i, CStr(keys(i)), x + px(16), y + px(96) + i * px(22), px(112), px(20), 9, False, CLR_MUTED_FG
        Lbl "lblVal" & i, "", x + px(136), y + px(96) + i * px(22), fw - px(120), px(20), 9, False, CLR_FG
    Next

    ' Email notices.
    y = y + h + px(12): h = px(112)
    Card "Notices", "Notices", x, y, w, h
    Toggle "tgEmail", "", x + px(20), y + px(52), px(36), px(20), "EmailChanged"
    Link "lbltgEmail", "Email notices", x + px(68), y + px(48), fw - px(52), px(22), "=ToggleSwitch([Form],""tgEmail"")", 10, CLR_FG
    Lbl "lblEmailSub", "", x + px(68), y + px(70), fw - px(52), px(32), 9, False, CLR_MUTED_FG   ' what goes where, or why off

    ' Your roles.
    x = PAGE_PAD + LEFT_W + px(12): w = RightW(): fw = w - px(32): y = px(84): h = px(284)
    Card "Roles", "Roles and permissions", x, y, w, h, , 2, 0
    Lbl "lblRolesNote", "Your roles decide the pages and actions you can use.", x + px(16), y + px(40), fw, px(18), _
        9, False, CLR_MUTED_FG
    GridHeaders "hdrRoles", RoleCols(), x + px(16), y + px(66), w:=fw
    Set c = SubF("sfRoles", "subGridProfileRoles", x + px(16), y + px(88), fw, h - px(104)): Anchor c, 2, 0
    EmptyState "emptyRoles", x + px(16), y + px(66), fw, "E8D7", "No roles to show", "Your roles show here once you're signed in."

    ' Colors.
    y = y + h + px(12): h = px(112)
    Card "Colors", "Colors", x, y, w, h, , 2, 0
    Lbl "lblColors", "Choose the palette the whole app uses on this computer: light, dark and more.", _
        x + px(16), y + px(40), fw, px(20), 9, False, CLR_MUTED_FG
    Set c = Btn("btnAppearance", "Appearance", x + px(16), y + px(72), px(120), px(28), "outline")
    c.ControlTipText = "Pick the app's colors"
    c.OnClick = "=AppAppearance()"

    ' Sign out.
    y = y + h + px(12): h = PAGE_H - y - px(16)
    Card "SignOut", "Sign out", x, y, w, h, , 2, 0
    Lbl "lblSignedIn", "", x + px(16), y + px(40), fw, px(36), 9, False, CLR_MUTED_FG   ' who is signed in (Form_Load)
    Set c = Btn("btnSignOut", "Sign out", x + px(16), y + px(80), px(120), px(28), "outline")
    c.ControlTipText = "Sign out so someone else can sign in on this copy"
    c.OnClick = "=AppSignOut()"
    SaveForm "pgProfile"
End Sub
