Option Compare Database
Option Explicit

' Access UI kit - the command palette: Ctrl+K (BuildShortcutKeys) opens a
' search box over the app's pages, the records the user opened lately and
' the sidebar's commands (Help, Appearance), as in a web app. Typing narrows
' the list (every word, anywhere in the name, ignoring case; names that
' start with the text come first), Up and Down move, Enter or a click opens
' the lit item, Esc or a click outside closes it. dlgPalette
' (BuildPaletteDialog, with dlgPalette.vba) opens a fifth of the way down
' the Access window and takes the focus while it is open.
'
' Pages and commands come from the sidebar ShellBuild built (ShellItems); an
' app without it lists its recent records only. A page records what the
' user opens with KitRecent, naming a public function that opens it again:
'   KitRecent "PO-1234 Pump overhaul", "OpenItem(1234)", "E8A5"
' The palette keeps the last PALETTE_RECENT, newest first, for this user on
' this PC (the registry, under APP_NAME).

Public Const PALETTE_ROWS As Long = 8      ' most items the list shows
Public Const PALETTE_W As Long = 560       ' px
Public Const PALETTE_ROW_H As Long = 36    ' px
Public Const PALETTE_RECENT As Long = 8    ' recent records kept
Private Const PAD As Long = 8              ' px, around the list
Private Const LIST_TOP As Long = 64        ' px: under the search box and its rule
Private Const FOOT_H As Long = 32          ' px: the keys hint
Private Const RECENT_GLYPH As String = "E81C"  ' History
Private Const TICK_MS As Long = 150        ' how often it checks it still has the focus

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" ( _
    ByVal hMonitor As LongPtr, lpmi As MONITORINFO) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4

Private m_items As Collection   ' every item: Array(caption, kind, glyph character, expression)
Private m_hits() As Long        ' the rows shown: indexes into m_items
Private m_count As Long         ' how many rows are shown
Private m_hot As Long           ' the lit row, -1 for none
Private m_closing As Boolean    ' dlgPalette is closing (Deactivate after a pick)

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' Build dlgPalette (with the app's other dialogs; copy kit\forms\dlgPalette.vba).
Public Sub BuildPaletteDialog()
    Dim c As Control, r As Long, w As Long, h As Long, y As Long, nm As Variant
    w = px(PALETTE_W)
    h = PaletteHeight(PALETTE_ROWS)
    NewForm "", 0, w, h
    With F
        .Caption = "Go to"
        .PopUp = True
        .Modal = False
        .BorderStyle = 0
        .AutoCenter = False
        .AutoResize = False
        .ControlBox = False
        .MinMaxButtons = 0
        .CloseButton = False
        .ShortcutMenu = False
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .Section(acDetail).BackColor = CLR_DIALOG
        .Section(acDetail).AlternateBackColor = CLR_DIALOG
        .Section(acDetail).OnMouseMove = "=PaletteHover([Form],-1)"
    End With
    Ev F, "OnLoad", "OnDeactivate", "OnTimer"
    Set c = SearchBox("txtSearch", px(12), px(12), w - px(24), "PaletteFilter", "Go to a page or a recent record", px(36))
    Ev c, "OnKeyDown"
    Rule "ruleSearch", 0, px(LIST_TOP - PAD \ 2), w, 0, CLR_BORDER
    For r = 0 To PALETTE_ROWS - 1
        y = px(LIST_TOP) + r * px(PALETTE_ROW_H)
        Set c = Box("boxRow" & r, px(PAD), y, w - 2 * px(PAD), px(PALETTE_ROW_H), CLR_OUT_HOVER)
        c.Visible = False
        ' The caption spans the row (its left margin clears the icon), so the
        ' mouse is over one of the row's parts anywhere in it; the kind sits
        ' over its right end.
        Set c = Lbl("lblCap" & r, "", px(PAD), y, w - 2 * px(PAD), px(PALETTE_ROW_H), 10, False, CLR_FG)
        c.LeftMargin = px(40): c.RightMargin = px(96)
        VCenter c, px(PALETTE_ROW_H), 10, True
        Set c = Lbl("lblIcon" & r, "", px(PAD + 10), y, px(20), px(PALETTE_ROW_H), 11, False, CLR_MUTED_FG)
        c.FontName = "Segoe MDL2 Assets": c.TextAlign = 2
        VCenter c, px(PALETTE_ROW_H), 11, True
        Set c = Lbl("lblSub" & r, "", w - px(PAD) - px(88), y, px(88), px(PALETTE_ROW_H), 9, False, CLR_MUTED_FG)
        c.TextAlign = 3: c.RightMargin = px(12)
        VCenter c, px(PALETTE_ROW_H), 9, True
        For Each nm In Array("boxRow", "lblIcon", "lblCap", "lblSub")
            F.Controls(nm & r).OnClick = "=PaletteClick([Form]," & r & ")"
            F.Controls(nm & r).OnMouseMove = "=PaletteHover([Form]," & r & ")"
        Next
    Next
    ' The lit row's accent bar (PaletteHover moves it), so the keyboard's row
    ' stands out from the mouse's hover tint.
    Set c = Box("barHot", px(PAD), px(LIST_TOP), px(3), px(PALETTE_ROW_H), CLR_PRIMARY)
    c.Visible = False
    Set c = Lbl("lblEmpty", "No matches", px(PAD), px(LIST_TOP), w - 2 * px(PAD), px(PALETTE_ROW_H), 10, False, CLR_MUTED_FG)
    c.LeftMargin = px(12): c.Visible = False
    VCenter c, px(PALETTE_ROW_H), 10, True
    y = h - px(FOOT_H)
    Rule "ruleFoot", 0, y, w, 0, CLR_BORDER
    Set c = Lbl("lblFoot", ChrW(&H2191) & ChrW(&H2193) & " to move     Enter to open     Esc to close", _
        0, y, w, px(FOOT_H), 8, False, CLR_MUTED_FG)
    c.LeftMargin = px(12)
    VCenter c, px(FOOT_H), 8, True
    Rule "ruleTop", 0, 15, w, 0, CLR_INPUT_BORDER
    Rule "ruleBottom", 0, h, w, 0, CLR_INPUT_BORDER
    Rule "ruleLeft", 15, 0, 0, h, CLR_INPUT_BORDER
    Rule "ruleRight", w, 0, 0, h, CLR_INPUT_BORDER
    SaveForm "dlgPalette"
End Sub

' The palette's height (twips) showing n rows.
Private Function PaletteHeight(ByVal n As Long) As Long
    PaletteHeight = px(LIST_TOP + n * PALETTE_ROW_H + PAD + FOOT_H)
End Function

' ---------------------------------------------------------------------------
' Run time
' ---------------------------------------------------------------------------

' Ctrl+K: open the palette, or close it when it is open. Does nothing in an
' app without dlgPalette.
Public Function KitPalette()
    On Error GoTo Fail
    If IsFormOpen("dlgPalette") Then
        PaletteClose Forms("dlgPalette")
    ElseIf PaletteBuilt() Then
        DoCmd.OpenForm "dlgPalette"
    End If
    Exit Function
Fail:
    HandleError "modKitPalette.KitPalette"
End Function

' Remember a record the user opened, for the palette's recent list: caption
' names it, expression calls a public function that opens it again
' ("OpenItem(12)", no "="), glyph is its MDL2 code (default History).
' Never raises: a page calls it on its way to showing the record.
Public Sub KitRecent(ByVal caption As String, ByVal expression As String, Optional ByVal glyph As String = RECENT_GLYPH)
    Dim list As Collection, i As Long, n As Long
    On Error GoTo Fail
    If Len(expression) = 0 Then Exit Sub
    Set list = RecentList()
    SaveSetting APP_NAME, PaletteRecentKey(), "1", Replace(caption, vbTab, " ") & vbTab & expression & vbTab & glyph
    n = 1
    For i = 1 To list.Count
        If n = PALETTE_RECENT Then Exit For
        If list(i)(3) <> expression Then
            n = n + 1
            SaveSetting APP_NAME, PaletteRecentKey(), CStr(n), list(i)(0) & vbTab & list(i)(3) & vbTab & list(i)(4)
        End If
    Next
    Exit Sub
Fail:
    LogError "modKitPalette.KitRecent", Left$(caption, 60)
End Sub

' The registry section for the signed-in person's recent records (KitUser),
' so people sharing a Windows account each see their own.
Public Function PaletteRecentKey() As String
    PaletteRecentKey = "Recent " & Replace(KitUser(), "\", "_")
End Function

' The recent records, newest first: Array(caption, "Recent", glyph character,
' expression, glyph code).
Private Function RecentList() As Collection
    Dim i As Long, p() As String
    Set RecentList = New Collection
    For i = 1 To PALETTE_RECENT
        p = Split(GetSetting(APP_NAME, PaletteRecentKey(), CStr(i), ""), vbTab)
        If UBound(p) = 2 Then
            If Len(p(1)) > 0 Then RecentList.Add Array(p(0), "Recent", GlyphChar(p(2)), p(1), p(2))
        End If
    Next
End Function

' An MDL2 code ("E81C") as its character; "" for none or a bad code.
Private Function GlyphChar(ByVal code As String) As String
    On Error Resume Next
    If Len(code) > 0 Then GlyphChar = ChrW(CLng("&H" & code))
End Function

' dlgPalette's Load: gather the items, place it, list them all.
Public Sub PaletteOpen(frm As Form)
    m_closing = False
    LoadItems
    frm.InsideWidth = px(PALETTE_W)
    frm.InsideHeight = PaletteHeight(PALETTE_ROWS)
    PlaceTop frm
    PaletteShow frm, ""
    frm.TimerInterval = TICK_MS
End Sub

' The items: the shell's pages, the recent records, then its commands.
Private Sub LoadItems()
    Dim sItems As Collection, v As Variant, kind As String, pass As Long
    Set m_items = New Collection
    Set sItems = New Collection
    If IsFormOpen(SHELL_FORM) Then Set sItems = ShellItems(Forms(SHELL_FORM))
    For pass = 0 To 2
        If pass = 1 Then
            For Each v In RecentList()
                m_items.Add v
            Next
        Else
            kind = IIf(pass = 0, "page", "link")
            For Each v In sItems
                If v(4) = kind Then m_items.Add Array(v(1), IIf(pass = 0, "Page", "Command"), v(2), v(3))
            Next
        End If
    Next
End Sub

' The search box's handler (dlgPalette.PaletteFilter): list the items whose
' names hold every word of text, those starting with it first, and light
' the first.
Public Sub PaletteShow(frm As Form, ByVal text As String)
    Dim i As Long, pass As Long, cap As String, r As Long, n As Long, capW As Long
    text = Trim$(text)
    ReDim m_hits(0 To PALETTE_ROWS - 1)
    m_count = 0
    For pass = 0 To 1
        For i = 1 To m_items.Count
            If m_count = PALETTE_ROWS Then Exit For
            cap = m_items(i)(0)
            If (InStr(1, cap, text, vbTextCompare) = 1) = (pass = 0) Then
                If PaletteMatch(cap, text) Then m_hits(m_count) = i: m_count = m_count + 1
            End If
        Next
    Next
    ' A caption stays on one line, with an ellipsis, clear of the icon and kind.
    capW = px(PALETTE_W) - 2 * px(PAD) - px(40) - px(96)
    For r = 0 To PALETTE_ROWS - 1
        If r < m_count Then
            frm.Controls("lblCap" & r).Caption = Replace(FitSingleLineText(m_items(m_hits(r))(0), capW, 10), "&", "&&")
            frm.Controls("lblIcon" & r).Caption = m_items(m_hits(r))(2)
            frm.Controls("lblSub" & r).Caption = m_items(m_hits(r))(1)
        End If
        frm.Controls("lblCap" & r).Visible = (r < m_count)
        frm.Controls("lblIcon" & r).Visible = (r < m_count)
        frm.Controls("lblSub" & r).Visible = (r < m_count)
        frm.Controls("boxRow" & r).Visible = False
    Next
    frm.Controls("lblEmpty").Visible = (m_count = 0)
    frm.Controls("barHot").Visible = False
    n = IIf(m_count = 0, 1, m_count)
    ' The window shrinks from the bottom: the controls move up before it does.
    frm.Controls("ruleFoot").Top = PaletteHeight(n) - px(FOOT_H)
    frm.Controls("lblFoot").Top = PaletteHeight(n) - px(FOOT_H)
    frm.Controls("ruleBottom").Top = PaletteHeight(n)
    frm.Controls("ruleLeft").Height = PaletteHeight(n)
    frm.Controls("ruleRight").Height = PaletteHeight(n)
    frm.InsideHeight = PaletteHeight(n)
    m_hot = -1
    If m_count > 0 Then PaletteHover frm, 0
End Sub

' True when caption holds every word of text (ignoring case).
Public Function PaletteMatch(ByVal caption As String, ByVal text As String) As Boolean
    Dim w As Variant
    For Each w In Split(Trim$(text), " ")
        If Len(w) > 0 Then
            If InStr(1, caption, w, vbTextCompare) = 0 Then Exit Function
        End If
    Next
    PaletteMatch = True
End Function

' The caption of row r as shown ("" past the last match), for the tests.
Public Function PaletteRowText(ByVal r As Long) As String
    If r >= 0 And r < m_count Then PaletteRowText = m_items(m_hits(r))(0)
End Function

' How many rows are shown, and which is lit (-1 for none).
Public Function PaletteRows() As Long
    PaletteRows = m_count
End Function

Public Function PaletteHot() As Long
    PaletteHot = m_hot
End Function

' A row's mouse move: light row r (-1 for none).
Public Function PaletteHover(frm As Form, ByVal r As Long)
    On Error GoTo Fail
    If r < 0 Or r >= m_count Or r = m_hot Then Exit Function
    If m_hot >= 0 Then frm.Controls("boxRow" & m_hot).Visible = False
    m_hot = r
    frm.Controls("boxRow" & r).Visible = True
    frm.Controls("barHot").Top = frm.Controls("boxRow" & r).Top
    frm.Controls("barHot").Visible = True
    Exit Function
Fail:
    HandleError "modKitPalette.PaletteHover", "row " & r
End Function

' A row's click: open it.
Public Function PaletteClick(frm As Form, ByVal r As Long)
    On Error GoTo Fail
    PaletteRun frm, r
    Exit Function
Fail:
    HandleError "modKitPalette.PaletteClick", "row " & r
End Function

' The search box's keys: Up and Down move the light (round the ends), Enter
' opens the lit row, Esc closes.
Public Sub PaletteKeyDown(frm As Form, KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If m_count > 0 Then PaletteHover frm, (m_hot + 1) Mod m_count
        Case vbKeyUp
            If m_count > 0 Then PaletteHover frm, (m_hot - 1 + m_count) Mod m_count
        Case vbKeyReturn
            If m_hot >= 0 Then PaletteRun frm, m_hot
        Case vbKeyEscape
            PaletteClose frm
        Case Else
            Exit Sub
    End Select
    KeyCode = 0
End Sub

' Close the palette, then do row r's item (go to the page, open the record,
' run the command).
Public Sub PaletteRun(frm As Form, ByVal r As Long)
    Dim expr As String
    If r < 0 Or r >= m_count Then Exit Sub
    expr = m_items(m_hits(r))(3)
    PaletteClose frm
    ' From the app window's Timer, not from inside this event: a page that
    ' saves its open record (a required field still empty) must do it there.
    If IsFormOpen(SHELL_FORM) Then
        ShellRunLater Forms(SHELL_FORM), expr
    Else
        Eval expr
    End If
End Sub

' dlgPalette's Deactivate: close it from its Timer (Access can't close a
' form in its own Deactivate).
Public Sub PaletteLostFocus(frm As Form)
    If Not m_closing Then frm.TimerInterval = 1
End Sub

' dlgPalette's Timer: close it once another form has the focus. A click on
' the app's window doesn't always raise the pop-up's Deactivate, so it checks.
Public Sub PaletteTick(frm As Form)
    Dim nm As String
    If m_closing Then Exit Sub
    On Error Resume Next                ' no active form: it lost the focus
    nm = Screen.ActiveForm.Name
    On Error GoTo 0
    If nm <> frm.Name Or frm.TimerInterval = 1 Then
        frm.TimerInterval = 0
        PaletteClose frm
    End If
End Sub

Public Sub PaletteClose(frm As Form)
    If m_closing Then Exit Sub
    m_closing = True
    DoCmd.Close acForm, frm.Name, acSaveNo
End Sub

' True when the app has a dlgPalette form (BuildPaletteDialog).
Private Function PaletteBuilt() As Boolean
    On Error Resume Next
    PaletteBuilt = (Len(CurrentProject.AllForms("dlgPalette").Name) > 0)
End Function

' Put the palette across the middle of the Access window, a fifth of the way
' down (it shrinks from the bottom as the list narrows), kept on its monitor.
Private Sub PlaceTop(frm As Form)
    Dim app As RECT, dlg As RECT, mi As MONITORINFO, hApp As LongPtr, x As Long, y As Long
    hApp = Application.hWndAccessApp
    If GetWindowRect(frm.hWnd, dlg) = 0 Then Exit Sub
    mi.cbSize = LenB(mi)
    If GetMonitorInfo(MonitorFromWindow(hApp, MONITOR_DEFAULTTONEAREST), mi) = 0 Then Exit Sub
    If IsIconic(hApp) <> 0 Then
        app = mi.rcWork
    ElseIf GetWindowRect(hApp, app) = 0 Then
        Exit Sub
    End If
    x = CenteredStart(app.Left, app.Right, dlg.Right - dlg.Left, mi.rcWork.Left, mi.rcWork.Right)
    y = app.Top + (app.Bottom - app.Top) \ 5
    If y + (dlg.Bottom - dlg.Top) > mi.rcWork.Bottom Then y = mi.rcWork.Bottom - (dlg.Bottom - dlg.Top)
    If y < mi.rcWork.Top Then y = mi.rcWork.Top
    SetWindowPos frm.hWnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
End Sub
