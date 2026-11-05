Option Compare Database
Option Explicit

' Access UI kit - the app shell: frmMain with a sidebar of pages on the left
' and the page host sfPage on the right. ShellBuild builds the form (an app
' adds its own controls, such as a classification banner, then saves it);
' frmMain's code-behind hands its events to the Shell* routines here, so
' every app swaps pages, guards unsaved work and collapses the sidebar the
' same way:
'   Private Sub Form_Load(): ShellLoad Me: End Sub
'   Private Sub Form_Timer(): ShellTimer Me: End Sub
'   Private Sub Form_Unload(Cancel As Integer): Cancel = Not ShellCanLeave(Me): End Sub
'   Public Property Get PageKey() As String: PageKey = ShellPageKey(): End Property
'   Public Sub QueueNav(ByVal key As String): ShellQueue Me, key: End Sub
'   Public Sub Reopen(ByVal pg As String): ShellReopen Me, pg: End Sub
'   Public Function CanLeavePage() As Boolean: CanLeavePage = ShellCanLeave(Me): End Function
' (each with the form's usual error handler). A sidebar item calls
' =ShellGo("key"), which asks frmMain's QueueNav, so an app's own NavTo can
' stay as a one-line call to ShellGo.
'
' Pages are "key|Caption|form|glyph" (glyph: a Segoe MDL2 Assets code in hex,
' E80F for Home); the first is the start page. Links, anchored to the bottom
' of the sidebar, are "key|Caption|glyph|=Expression()" (Help, Appearance,
' Sign out). A collapsible sidebar folds to a 60px rail of glyphs, and
' remembers it for the next time the app opens.
'
' Optional hooks in frmMain's code-behind (the shell calls them when they
' exist):
'   Public Function RoutePage(ByVal key As String) As String
'       the page to show instead (a page the user can't open sends them to
'       another); return key to show it.
'   Public Function PageForm(ByVal key As String) As String
'       the form for a key that isn't in the sidebar (a detail page).
'   Public Sub PageShown(ByVal key As String)
'       after each page is shown (a count in the sidebar, a tab to restore).
' Give each hook the form's usual error handler: CallByName can't tell an
' error 438 or 2465 escaping a hook from a missing hook, so it would read as none
' (and RoutePage's redirect would be skipped).
' A page with unsaved work it can't save yet has a Public Function
' CanLeave() As Boolean that saves it or says why not (ShellCanLeave).

Public Const SHELL_FORM As String = "frmMain"
Public Const SHELL_W As Long = 3000         ' the sidebar, 200px (twips)
Public Const SHELL_RAIL_W As Long = 900     ' collapsed, 60px (twips)

Private Const ITEM_H As Long = 34           ' a sidebar item's height (px)
Private Const ITEM_STEP As Long = 38        ' from one item to the next (px)

Private m_page As String        ' the page shown
Private m_pending As String     ' the page QueueNav asked for ("*reopen": reopen the form)
Private m_reopening As Boolean  ' frmMain is closing to reopen (ShellReopening)
Private m_run As String         ' the expression ShellRunLater asked for

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' Start frmMain (NewForm) with its sidebar and an empty sfPage (ShellLoad
' shows the start page). pageW x pageH is the page area; top and bottom
' leave room above and below the sidebar and page (a banner). Raises error 5
' for no pages, a bad page or link, or a collapsible sidebar item without a
' glyph. The caller adds its own controls, then SaveForm SHELL_FORM.
Public Sub ShellBuild(ByVal title As String, ByVal subtitle As String, pages As Variant, links As Variant, _
        ByVal pageW As Long, ByVal pageH As Long, Optional ByVal top As Long = 0, Optional ByVal bottom As Long = 0, _
        Optional ByVal collapsible As Boolean = True)
    Dim c As Control, i As Long, p() As String, y As Long, keys As String, brandW As Long
    ShellCheck pages, links, collapsible
    NewForm "", 0, SHELL_W + pageW, top + pageH + bottom
    Ev F, "OnLoad", "OnTimer", "OnResize", "OnUnload"
    F.Caption = title
    Set c = Box("boxSide", 0, top, SHELL_W, pageH, CLR_SIDEBAR): Anchor c, 0, 2
    brandW = IIf(collapsible, px(132), px(160))
    Set c = Lbl("lblBrand", title, px(20), top + px(20), brandW, px(26), 13, True, CLR_WHITE): c.Tag = "shell:full"
    Set c = Lbl("lblBrandSub", subtitle, px(20), top + px(46), brandW, px(18), 8, False, CLR_NAV_MUTED): c.Tag = "shell:full"
    Set c = Rule("ruleSide", px(16), top + px(76), px(168), 0, CLR_SIDEBAR_LINE)
    If collapsible Then
        Set c = Link("lnkCollapse", ChrW(&HE700), px(156), top + px(20), px(32), px(28), "=ShellToggle()", 12, CLR_NAV_FG)
        c.FontName = "Segoe MDL2 Assets": c.FontWeight = 400: c.TextAlign = 2
        c.ControlTipText = "Collapse the sidebar"
        VCenter c, px(28), 12, True
    End If
    y = top + px(90)
    For i = 0 To UBound(pages)
        p = Split(pages(i), "|")
        ShellItem p(0), p(1), p(3), y, "=ShellGo(""" & p(0) & """)", "shell|" & p(2)
        keys = keys & IIf(Len(keys) > 0, ",", "") & p(0)
        y = y + px(ITEM_STEP)
    Next
    ' The pages in order, and where the first one sits, for ShellShowItem.
    F.Controls("boxSide").Tag = "shell:" & (top + px(90)) & ":" & keys
    If IsArray(links) Then
        y = top + pageH - px(12) - (UBound(links) - LBound(links) + 1) * px(ITEM_STEP) + px(ITEM_STEP - ITEM_H)
        For i = LBound(links) To UBound(links)
            p = Split(links(i), "|")
            ShellItem p(0), p(1), p(2), y, p(3), "shell:link"
            Anchor F.Controls("nb_" & p(0)), 0, 1
            If Len(p(2)) > 0 Then Anchor F.Controls("ni_" & p(0)), 0, 1
            Anchor F.Controls("nt_" & p(0)), 0, 1
            y = y + px(ITEM_STEP)
        Next
    End If
    Set c = SubF("sfPage", "", SHELL_W, top, pageW, pageH): Anchor c, 2, 2
End Sub

' One sidebar item: the highlight nb_<key> (behind), the glyph ni_<key> and
' the caption nt_<key>, all clicking to the same action. Without a glyph the
' caption starts where the glyph would be.
Private Sub ShellItem(ByVal key As String, ByVal cap As String, ByVal glyph As String, ByVal y As Long, _
        ByVal action As String, ByVal tag As String)
    Dim c As Control, x As Long
    Set c = Link("nb_" & key, "", px(12), y, px(176), px(ITEM_H), action, 10, CLR_NAV_FG)
    c.BackColor = CLR_NAV_ACTIVE_BG
    c.Tag = tag
    c.ControlTipText = cap
    x = px(24)
    If Len(glyph) > 0 Then
        Set c = Link("ni_" & key, ChrW(CLng("&H" & glyph)), px(20), y, px(24), px(ITEM_H), action, 11, CLR_NAV_FG)
        c.FontName = "Segoe MDL2 Assets": c.FontWeight = 400: c.TextAlign = 2
        c.ControlTipText = cap
        VCenter c, px(ITEM_H), 11, True
        x = px(52)
    End If
    Set c = Link("nt_" & key, cap, x, y, px(188) - x, px(ITEM_H), action, 10, CLR_NAV_FG)
    c.ControlTipText = cap
    VCenter c, px(ITEM_H), 10, True
End Sub

' Raise error 5 for no pages, a page or link missing a part, or (when the
' sidebar collapses to glyphs) an item without one.
Private Sub ShellCheck(pages As Variant, links As Variant, ByVal collapsible As Boolean)
    Dim i As Long, p() As String
    If Not IsArray(pages) Then Err.Raise 5, , "ShellBuild has no pages"
    If UBound(pages) < LBound(pages) Then Err.Raise 5, , "ShellBuild has no pages"
    For i = LBound(pages) To UBound(pages)
        p = Split(pages(i), "|")
        If UBound(p) <> 3 Then Err.Raise 5, , "ShellBuild page is not key|Caption|form|glyph: " & pages(i)
        If Len(p(0)) = 0 Or Len(p(2)) = 0 Then Err.Raise 5, , "ShellBuild page has no key or form: " & pages(i)
        If collapsible And Len(p(3)) = 0 Then Err.Raise 5, , "ShellBuild page has no glyph for the rail: " & pages(i)
    Next
    If Not IsArray(links) Then Exit Sub
    For i = LBound(links) To UBound(links)
        p = Split(links(i), "|")
        If UBound(p) <> 3 Then Err.Raise 5, , "ShellBuild link is not key|Caption|glyph|=Expression(): " & links(i)
        If Len(p(0)) = 0 Or Left$(p(3), 1) <> "=" Then Err.Raise 5, , "ShellBuild link has no key or action: " & links(i)
        If collapsible And Len(p(2)) = 0 Then Err.Raise 5, , "ShellBuild link has no glyph for the rail: " & links(i)
    Next
End Sub

' ---------------------------------------------------------------------------
' Run time (frmMain's events)
' ---------------------------------------------------------------------------

' frmMain's Load: the sidebar as the user left it, then the page TempVars
' StartPage names (a Reopen), else start, else the first page.
Public Sub ShellLoad(frm As Form, Optional ByVal start As String = "")
    m_reopening = False
    m_pending = ""
    If HasControl(frm, "lnkCollapse") Then
        ShellCollapse frm, (GetSetting(APP_NAME, "UI", "SidebarCollapsed", "0") = "1"), False
    End If
    If Len(Nz(TempVars("StartPage"), "")) > 0 Then start = TempVars("StartPage")
    TempVars("StartPage") = Null
    ShellShow frm, start
End Sub

' A sidebar item's click: go to page key (frmMain's QueueNav, so the app's
' own checks run).
Public Function ShellGo(ByVal key As String)
    On Error GoTo Err_Handler
    If IsFormOpen(SHELL_FORM) Then Forms(SHELL_FORM).QueueNav key
    Exit Function
Err_Handler:
    HandleError "modKitShell.ShellGo", key
End Function

' Go to page key once the calling event has finished (frmMain's Timer), so a
' page is never swapped while one of its own events is running. Not while
' the current page refuses to be left.
Public Sub ShellQueue(frm As Form, ByVal key As String)
    If Not ShellCanLeave(frm) Then Exit Sub
    m_pending = key
    frm.TimerInterval = 10
End Sub

' Close and reopen frmMain on page pg (after a palette switch).
Public Sub ShellReopen(frm As Form, ByVal pg As String)
    TempVars("StartPage") = pg
    ShellQueue frm, "*reopen"
End Sub

' Run expression ("OpenItem(12)", no "=") from frmMain's Timer, once the
' calling event has finished: a pop-up that has just closed (the command
' palette) hands its pick to the page this way, so the page's own save and
' its required-field message run as they would from a click.
Public Sub ShellRunLater(frm As Form, ByVal expression As String)
    m_run = expression
    frm.TimerInterval = 10
End Sub

' frmMain's Timer: the expression ShellRunLater asked for, then the page
' ShellQueue asked for (the expression may have asked for one).
Public Sub ShellTimer(frm As Form)
    Dim key As String, expr As String
    frm.TimerInterval = 0
    expr = m_run
    m_run = ""
    If Len(expr) > 0 Then
        Eval expr
        frm.TimerInterval = 0
    End If
    key = m_pending
    m_pending = ""
    If key = "*reopen" Then
        m_reopening = True
        DoCmd.Close acForm, frm.Name, acSaveNo
        DoCmd.OpenForm SHELL_FORM
    ElseIf Len(key) > 0 Then
        ShellShow frm, key
    End If
End Sub

' True while frmMain closes to reopen (its Close shouldn't sign the user
' out or end the session).
Public Function ShellReopening() As Boolean
    ShellReopening = m_reopening
End Function

' The page shown (frmMain's PageKey).
Public Function ShellPageKey() As String
    ShellPageKey = m_page
End Function

' True when the page shown can be swapped or closed: a page with unsaved
' work it can't save yet has a public CanLeave that saves it or says why not
' and returns False. Pages without CanLeave can always be left. The sidebar
' items are labels, so clicking them never takes the focus out of the page,
' which is what would otherwise make Access save it. A save the page's form
' refused with its own message (KitRefuseSave) keeps the page, even when the
' error Access raises for it is 2465, which a page without CanLeave raises too.
Public Function ShellCanLeave(frm As Form) As Boolean
    Dim ok As Boolean, errNo As Long, errDesc As String, errSrc As String
    ShellCanLeave = True
    If Len(frm.sfPage.SourceObject) = 0 Then Exit Function
    KitTakeRefused
    On Error Resume Next
    ok = frm.sfPage.Form.CanLeave()
    errNo = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    On Error GoTo 0
    If KitTakeRefused() Then ShellCanLeave = False: Exit Function
    Select Case errNo
        Case 0: ShellCanLeave = ok
        Case 438, 2465          ' the page has no CanLeave
        Case Else: Err.Raise errNo, errSrc, errDesc
    End Select
End Function

' Show page key now (ShellQueue waits for the calling event; use that). The
' app's RoutePage may send it elsewhere; a key with no page shows the first.
Public Sub ShellShow(frm As Form, ByVal key As String)
    Dim formName As String, routed As String
    If Hook(frm, "RoutePage", key, routed) Then key = routed
    formName = ShellPageForm(frm, key)
    If Len(formName) = 0 Then
        key = Split(ShellKeys(frm), ",")(0)
        formName = ShellPageForm(frm, key)
    End If
    m_page = key
    If frm.sfPage.SourceObject <> formName Then frm.sfPage.SourceObject = formName
    ShellRefreshNav frm
    Hook frm, "PageShown", key, routed
End Sub

' The form for page key: its sidebar item's, else the app's PageForm, else "".
Private Function ShellPageForm(frm As Form, ByVal key As String) As String
    Dim t As String, nm As String
    If Len(key) = 0 Then Exit Function
    If HasControl(frm, "nb_" & key) Then t = frm.Controls("nb_" & key).Tag
    If Left$(t, 6) = "shell|" Then
        ShellPageForm = Mid$(t, 7)
    ElseIf Hook(frm, "PageForm", key, nm) Then
        ShellPageForm = nm
    End If
End Function

' Light the page shown in the sidebar (also after a palette change: the
' colors come from Th).
Public Sub ShellRefreshNav(frm As Form)
    Dim k As Variant, lit As Boolean
    For Each k In Split(ShellKeys(frm), ",")
        lit = (k = m_page)
        With frm.Controls("nb_" & k)
            .BackStyle = IIf(lit, 1, 0)
            .BackColor = Th(CLR_NAV_ACTIVE_BG)
        End With
        If HasControl(frm, "ni_" & k) Then frm.Controls("ni_" & k).ForeColor = IIf(lit, Th(CLR_NAV_ACTIVE_FG), Th(CLR_NAV_FG))
        With frm.Controls("nt_" & k)
            .ForeColor = IIf(lit, Th(CLR_NAV_ACTIVE_FG), Th(CLR_NAV_FG))
            .FontWeight = IIf(lit, 600, 400)
        End With
    Next
End Sub

' Show or hide page key's sidebar item (a page only admins see); the items
' below move up or down so there is no gap.
Public Sub ShellShowItem(frm As Form, ByVal key As String, ByVal visible As Boolean)
    Dim k As Variant, y As Long, p() As String
    For Each k In Array("nb_", "ni_", "nt_")
        If HasControl(frm, k & key) Then frm.Controls(k & key).Visible = visible
    Next
    p = Split(frm.Controls("boxSide").Tag, ":")
    y = CLng(p(1))
    For Each k In Split(p(2), ",")
        If frm.Controls("nb_" & k).Visible Then
            frm.Controls("nb_" & k).Top = y
            If HasControl(frm, "ni_" & k) Then frm.Controls("ni_" & k).Top = y
            frm.Controls("nt_" & k).Top = y
            y = y + px(ITEM_STEP)
        End If
    Next
End Sub

' The collapse button: fold the sidebar to its rail of glyphs, or open it.
Public Function ShellToggle()
    On Error GoTo Err_Handler
    If Not IsFormOpen(SHELL_FORM) Then Exit Function
    ShellCollapse Forms(SHELL_FORM), Forms(SHELL_FORM).Controls("boxSide").Width > SHELL_RAIL_W, True
    Exit Function
Err_Handler:
    HandleError "modKitShell.ShellToggle"
End Function

' Fold the sidebar to its 60px rail (collapsed) or open it to 200px; save
' the choice when remember. Captions and the brand hide on the rail, the
' highlights narrow to 44px around the centered glyphs, and the page takes
' the width the sidebar gives up.
Public Sub ShellCollapse(frm As Form, ByVal collapsed As Boolean, Optional ByVal remember As Boolean = True)
    Dim c As Control, w As Long, d As Long, isItem As Boolean
    w = IIf(collapsed, SHELL_RAIL_W, SHELL_W)
    d = frm.Controls("boxSide").Width - w         ' how much wider the page gets
    If remember Then SaveSetting APP_NAME, "UI", "SidebarCollapsed", IIf(collapsed, "1", "0")
    If d = 0 Then Exit Sub
    frm.Painting = False
    On Error GoTo Fail
    ' Widening: move the page first; narrowing: shrink it first, so it never
    ' runs past the form.
    If d > 0 Then frm.sfPage.Left = w: frm.sfPage.Width = frm.sfPage.Width + d
    If d < 0 Then frm.sfPage.Width = frm.sfPage.Width + d: frm.sfPage.Left = w
    frm.Controls("boxSide").Width = w
    For Each c In frm.Controls
        isItem = (Left$(c.Tag, 5) = "shell" And Left$(c.Name, 3) = "nb_")
        If c.Tag = "shell:full" Or Left$(c.Name, 3) = "nt_" Then
            c.Visible = Not collapsed
        ElseIf isItem Then
            c.Left = IIf(collapsed, px(8), px(12))
            c.Width = IIf(collapsed, px(44), px(176))
        ElseIf Left$(c.Name, 3) = "ni_" Then
            c.Left = IIf(collapsed, px(8), px(20))
            c.Width = IIf(collapsed, px(44), px(24))
        End If
    Next
    With frm.Controls("ruleSide")
        .Left = IIf(collapsed, px(12), px(16))
        .Width = IIf(collapsed, px(36), px(168))
    End With
    With frm.Controls("lnkCollapse")
        .Left = IIf(collapsed, px(14), px(156))
        .ControlTipText = IIf(collapsed, "Open the sidebar", "Collapse the sidebar")
    End With
    frm.Painting = True
    GridScrollAll frm
    Exit Sub
Fail:
    frm.Painting = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' The sidebar's items the user can see, pages then links, each
' Array(key, caption, glyph character or "", action expression without its
' "=", "page" or "link"), for the command palette (KitPalette). Empty for a
' form ShellBuild didn't build.
Public Function ShellItems(frm As Form) As Collection
    Dim c As Control, k As String, g As String
    Set ShellItems = New Collection
    If Not HasControl(frm, "boxSide") Then Exit Function
    If Left$(Nz(frm.Controls("boxSide").Tag, ""), 6) <> "shell:" Then Exit Function
    For Each c In frm.Controls
        If Left$(c.Name, 3) = "nb_" And Left$(c.Tag, 5) = "shell" And c.Visible Then
            k = Mid$(c.Name, 4)
            g = ""
            If HasControl(frm, "ni_" & k) Then g = frm.Controls("ni_" & k).Caption
            ShellItems.Add Array(k, frm.Controls("nt_" & k).Caption, g, Mid$(c.OnClick, 2), _
                IIf(c.Tag = "shell:link", "link", "page"))
        End If
    Next
End Function

' The sidebar's page keys in order, "a,b,c".
Private Function ShellKeys(frm As Form) As String
    ShellKeys = Split(frm.Controls("boxSide").Tag, ":")(2)
End Function

' True when frm has a control named nm.
Private Function HasControl(frm As Form, ByVal nm As String) As Boolean
    Dim c As Control
    On Error Resume Next
    Set c = frm.Controls(nm)
    HasControl = (Err.Number = 0)
End Function

' Call frmMain's optional hook nm with key; True, with its result, when the
' form has it. Errors inside the hook go to the caller.
Private Function Hook(frm As Form, ByVal nm As String, ByVal key As String, result As String) As Boolean
    Dim v As Variant, errNo As Long, errDesc As String, errSrc As String
    On Error Resume Next
    v = CallByName(frm, nm, VbMethod, key)
    errNo = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    On Error GoTo 0
    Select Case errNo
        Case 0
            Hook = True
            result = Nz(v, "")
        Case 438, 2465          ' the form has no such hook (a form raises 2465)
        Case Else: Err.Raise errNo, errSrc, errDesc
    End Select
End Function
