Option Compare Database
Option Explicit

' Access UI kit - dropdown action menu (a web "dropdown menu" or kebab menu):
' KitMenu opens a short list of actions under the button that was clicked and
' returns the key of the one the user picks. The menu (dlgMenu, built by
' BuildMenuDialog) is a borderless pop-up with a hairline edge, sized to its
' items, placed under the button (above it when there's no room below, and
' kept on the button's monitor). A row lights up under the mouse; Up, Down,
' Home, End and Enter work from the keyboard. A click anywhere else, Esc or
' Tab closes it with nothing picked.
'
' Items are "key|Caption" or "key|Caption|glyph", glyph a Segoe MDL2 Assets
' code point in hex (E70F is Edit, E74D Delete). "-" is a separator. A caption
' starting with "!" is destructive (CLR_DESTRUCTIVE, for Delete); a key
' starting with "~" is disabled (CLR_MUTED_FG, can't be picked).
'
' The module also has KitPopover, a short explanation shown by a control in
' the same kind of pop-up (see Popover below).

Public Const MENU_ROWS As Long = 10        ' most items a menu takes, separators included
Public Const MENU_ROW_H As Long = 32       ' an item's height (px)
Public Const MENU_SEP_H As Long = 9        ' a separator's height (px)
Public Const MENU_PAD As Long = 4          ' inside the edge, around the rows (px)
Public Const MENU_MIN_W As Long = 160      ' narrowest menu (px)
Public Const MENU_MAX_W As Long = 320      ' widest menu (px), its build width
Private Const MENU_GAP As Long = 4         ' between the button and the menu (px)
Private Const MENU_TICK_MS As Long = 150   ' how often an open menu checks it still has the focus
Private Const MENU_GRACE_MS As Long = 500  ' after opening, a menu that lost the focus takes it back

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function MonitorFromRect Lib "user32" (lprc As RECT, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" ( _
    ByVal hMonitor As LongPtr, lpmi As MONITORINFO) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, _
    ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, _
    ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, _
    ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, _
    ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

Private Const LOGPIXELSY As Long = 90
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const SWP_NOZORDER As Long = &H4

Private m_items As Variant      ' the open menu's items
Private m_pick As String        ' the key picked, "" for none
Private m_hot As Long           ' the lit row, -1 for none
Private m_anchor As RECT        ' the button, screen pixels
Private m_hAnchor As LongPtr    ' the button's form window (for its monitor), 0 at the pointer
' Per pop-up, so a popover closing doesn't stop the menu closing: index 0
' for dlgMenu, 1 for dlgPopover (PopupSlot).
Private m_closing(0 To 1) As Boolean
Private m_lost(0 To 1) As Boolean       ' its Deactivate ran
Private m_openedAt(0 To 1) As Single    ' Timer() when it opened
' The popover's (KitPopover).
Public Const POPOVER_W As Long = 280       ' px
Private Const POPOVER_LINES As Long = 8
Public Const POPOVER_LINE_H As Long = 255   ' one line of 9pt text, 17px (twips)
Private m_popTitle As String
Private m_popText As String
' Per pop-up (PopupSlot): the control it is by and that control's
' ControlTipText, cleared while it shows. Access's own tooltip on the anchor
' would otherwise sit over the pop-up's top edge while the pointer stays on it.
Private m_tipAnchor(0 To 1) As Control
Private m_tip(0 To 1) As String

' The searchable picker (PickBox, below).
Public Const PICK_ROWS As Long = 8         ' most matches the list shows
Public Const PICK_ROW_H As Long = 32       ' a row's height (px)
Public Const PICK_MIN_W As Long = 240      ' narrowest list (px)
Public Const PICK_MAX_W As Long = 480      ' widest list (px), its build width
Private Const PICK_TICK_MS As Long = 150   ' how often the list checks the box still has the focus
Private Const PICK_TAG As String = "pick|" ' a PickBox's Tag: this, then rowSql

Private m_pickBox As Control    ' the box the open list belongs to
Private m_pickKeys As Collection ' the listed rows' keys, first row first
Private m_pickText As Collection ' and their second columns, what the box shows
Private m_pickHot As Long       ' the lit row, -1 for none
Private m_pickEscAt As Single   ' Timer when an Esc cancelled the list's query (PickEscaped)

' The same menu at the mouse pointer, for a right-click (a table row's
' actions: subGrid calls the host's GridRightClick, which calls this).
'   Select Case KitMenuAt("edit|Edit|E70F", "-", "del|!Delete|E74D")
Public Function KitMenuAt(ParamArray items() As Variant) As String
    Dim list As Variant
    list = items
    If Not MenuOpenAt(list) Then Exit Function
    Do While IsFormOpen("dlgMenu")
        DoEvents
        Sleep 15
    Loop
    KitMenuAt = m_pick
End Function

' Show a menu of items under button anchor and wait for the user: returns the
' picked item's key, or "" when the menu was closed with nothing picked.
' Raises error 5 for no items, more than MENU_ROWS, or a bad item. Without
' dlgMenu (not built) it returns "" and logs why.
'   Select Case KitMenu(Me.btnMore, "edit|Edit|E70F", "copy|Duplicate|E8C8", "-", "del|!Delete|E74D")
Public Function KitMenu(anchor As Control, ParamArray items() As Variant) As String
    Dim list As Variant
    list = items
    If Not MenuOpen(anchor, list) Then Exit Function
    Do While IsFormOpen("dlgMenu")
        DoEvents
        Sleep 15
    Loop
    KitMenu = m_pick
End Function

' Open the menu without waiting (KitMenu waits; tests drive it this way).
' Returns False, logging why, when dlgMenu isn't built, and False when a menu
' is already open (a second click while one waits), leaving that one as it is.
Public Function MenuOpen(anchor As Control, items As Variant) As Boolean
    If IsFormOpen("dlgMenu") Then Exit Function
    MenuCheck items
    AnchorRect anchor
    HideTip 0, anchor
    MenuOpen = MenuShowItems(items)
    If Not MenuOpen Then RestoreTip 0
End Function

' Open the menu at the mouse pointer without waiting (KitMenuAt waits).
Public Function MenuOpenAt(items As Variant) As Boolean
    Dim pt As POINTAPI
    If IsFormOpen("dlgMenu") Then Exit Function     ' as MenuOpen
    MenuCheck items
    GetCursorPos pt
    m_anchor.Left = pt.x: m_anchor.Right = pt.x
    m_anchor.Top = pt.y: m_anchor.Bottom = pt.y
    m_hAnchor = 0                       ' Place uses the pointer's monitor
    MenuOpenAt = MenuShowItems(items)
End Function

' Raise error 5 for no items, more than MENU_ROWS, or a bad item.
Private Sub MenuCheck(items As Variant)
    Dim i As Long, p() As String
    If Not IsArray(items) Then Err.Raise 5, , "KitMenu has no items"
    If UBound(items) < LBound(items) Then Err.Raise 5, , "KitMenu has no items"
    If UBound(items) - LBound(items) + 1 > MENU_ROWS Then Err.Raise 5, , "KitMenu takes at most " & MENU_ROWS & " items"
    For i = LBound(items) To UBound(items)
        If CStr(items(i)) <> "-" Then
            p = Split(CStr(items(i)), "|")
            If UBound(p) < 1 Or UBound(p) > 2 Then Err.Raise 5, , "KitMenu item is not key|Caption[|glyph]: " & items(i)
            If Len(p(0)) = 0 Then Err.Raise 5, , "KitMenu item has no key: " & items(i)
        End If
    Next
End Sub

' Open dlgMenu on items at the anchor already noted. False, logging why,
' when dlgMenu isn't built.
Private Function MenuShowItems(items As Variant) As Boolean
    m_items = items
    m_pick = ""
    m_hot = -1
    m_closing(0) = False
    m_lost(0) = False
    On Error GoTo NotBuilt
    DoCmd.OpenForm "dlgMenu"
    MenuShowItems = True
    Exit Function
NotBuilt:
    LogError "modMenu.MenuOpen", "dlgMenu could not open: " & Err.Description
End Function

' The key the last menu returned ("" for none).
Public Function MenuResult() As String
    MenuResult = m_pick
End Function

' dlgMenu's Load: fill the rows, size the window to them and place it.
Public Sub MenuShow(frm As Form)
    Dim i As Long, r As Long, y As Long, w As Long, p() As String, cap As String, ok As Boolean
    Dim capW As Long, n As Long
    n = UBound(m_items) - LBound(m_items) + 1
    ' As wide as the widest caption, within MENU_MIN_W and MENU_MAX_W.
    w = px(MENU_MIN_W)
    For i = 0 To n - 1
        If ItemText(i) <> "-" Then
            capW = px(40) + MeasureText(ItemCaption(i), 9, 400) + px(24) + 2 * px(MENU_PAD)
            If capW > w Then w = capW
        End If
    Next
    If w > px(MENU_MAX_W) Then w = px(MENU_MAX_W)
    y = px(MENU_PAD)
    For r = 0 To MENU_ROWS - 1
        If r < n Then
            If ItemText(r) = "-" Then
                RowParts frm, r, False
                With frm.Controls("ruleSep" & r)
                    .Top = y + px(MENU_SEP_H \ 2) + 15: .Width = w - 2 * px(MENU_PAD): .Visible = True
                End With
                y = y + px(MENU_SEP_H)
            Else
                p = Split(ItemText(r), "|")
                cap = ItemCaption(r)
                frm.Controls("ruleSep" & r).Visible = False
                frm.Controls("boxRow" & r).Top = y
                frm.Controls("boxRow" & r).Width = w - 2 * px(MENU_PAD)
                frm.Controls("lblIcon" & r).Top = y
                frm.Controls("lblCap" & r).Top = y
                frm.Controls("lblCap" & r).Width = w - 2 * px(MENU_PAD)
                frm.Controls("lblCap" & r).Caption = Replace(cap, "&", "&&")
                frm.Controls("lblIcon" & r).Caption = ""
                If UBound(p) >= 2 Then
                    If Len(p(2)) > 0 Then frm.Controls("lblIcon" & r).Caption = ChrW(CLng("&H" & p(2)))
                End If
                RowColors frm, r
                RowParts frm, r, True
                frm.Controls("boxRow" & r).Visible = False
                y = y + px(MENU_ROW_H)
            End If
        Else
            RowParts frm, r, False
            frm.Controls("ruleSep" & r).Visible = False
        End If
    Next
    y = y + px(MENU_PAD)
    ' Hairline edges: each line paints 1px up and left of where it is placed.
    frm.Controls("ruleTop").Width = w
    frm.Controls("ruleBottom").Top = y: frm.Controls("ruleBottom").Width = w
    frm.Controls("ruleLeft").Height = y
    frm.Controls("ruleRight").Left = w: frm.Controls("ruleRight").Height = y
    frm.InsideWidth = w
    frm.InsideHeight = y
    Place frm
End Sub

' Where a menu w x h (pixels) goes for a button at aL..aR, aT..aB on a screen
' work area wL..wR, wT..wB: under the button, left edges lined up, gap pixels
' below it; right edges lined up when it would run off the right; above the
' button when it would run off the bottom; never off the top or left.
Public Sub MenuSpot(ByVal aL As Long, ByVal aT As Long, ByVal aR As Long, ByVal aB As Long, ByVal w As Long, ByVal h As Long, _
        ByVal wL As Long, ByVal wT As Long, ByVal wR As Long, ByVal wB As Long, ByVal gap As Long, x As Long, y As Long)
    x = aL
    If x + w > wR Then x = aR - w
    If x < wL Then x = wL
    y = aB + gap
    If y + h > wB Then y = aT - gap - h
    If y < wT Then y = wT
End Sub

' A row's click (set by BuildMenuDialog): pick its item, unless disabled.
Public Function MenuPick(frm As Form, ByVal r As Long)
    On Error GoTo Fail
    If Not Pickable(r) Then Exit Function
    m_pick = Split(ItemText(r), "|")(0)
    MenuClose frm
    Exit Function
Fail:
    HandleError "modMenu.MenuPick", "row " & r
End Function

' A row's mouse move (set by BuildMenuDialog): light row r (-1 for none).
Public Function MenuHover(frm As Form, ByVal r As Long)
    On Error GoTo Fail
    If r = m_hot Then Exit Function
    If r >= 0 Then If Not Pickable(r) Then r = -1
    If m_hot >= 0 Then frm.Controls("boxRow" & m_hot).Visible = False
    m_hot = r
    If r >= 0 Then frm.Controls("boxRow" & r).Visible = True
    Exit Function
Fail:
    LogError "modMenu.MenuHover", "row " & r
End Function

' dlgMenu's KeyDown: Up, Down, Home and End move the lit row (skipping
' separators and disabled items), Enter or Space picks it, Esc and Tab close.
Public Sub MenuKey(frm As Form, KeyCode As Integer)
    Dim n As Long, r As Long, stp As Long
    n = UBound(m_items) - LBound(m_items) + 1
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp, vbKeyHome, vbKeyEnd
            Select Case KeyCode
                Case vbKeyDown: stp = 1: r = m_hot + 1
                Case vbKeyUp: stp = -1: r = IIf(m_hot < 0, n - 1, m_hot - 1)
                Case vbKeyHome: stp = 1: r = 0
                Case vbKeyEnd: stp = -1: r = n - 1
            End Select
            Do While r >= 0 And r < n
                If Pickable(r) Then MenuHover frm, r: Exit Do
                r = r + stp
            Loop
        Case vbKeyReturn, vbKeySpace
            If m_hot >= 0 Then MenuPick frm, m_hot
        Case vbKeyEscape, vbKeyTab
            MenuClose frm
        Case Else
            Exit Sub
    End Select
    KeyCode = 0
End Sub

' Close the menu or popover (once: its Deactivate also asks, as it closes).
' The anchor's tooltip comes back here too, for an app's dlgMenu built before
' it had a Close event.
Public Function MenuClose(frm As Form)
    If m_closing(PopupSlot(frm)) Then Exit Function
    m_closing(PopupSlot(frm)) = True
    RestoreTip PopupSlot(frm)
    DoCmd.Close acForm, frm.Name
End Function

' The Deactivate of dlgMenu or dlgPopover: a click outside it. Closing during
' the event is refused, so the form's timer closes it a moment later.
Public Sub MenuLostFocus(frm As Form)
    m_lost(PopupSlot(frm)) = True
    If Not m_closing(PopupSlot(frm)) Then frm.TimerInterval = 1
End Sub

' The Load of dlgMenu or dlgPopover, once it has the focus: start checking
' that it keeps it. Access doesn't always raise Deactivate for a click
' outside the pop-up, and a menu left open would keep KitMenu waiting.
Public Sub MenuStarted(frm As Form)
    m_openedAt(PopupSlot(frm)) = Timer
    frm.TimerInterval = MENU_TICK_MS
End Sub

' The timer of dlgMenu or dlgPopover: close it when its Deactivate ran or
' another form has the focus. Just after it opens (the second click of a
' double-click on its button), it takes the focus back instead, so Esc
' still reaches it.
Public Sub MenuTick(frm As Form)
    Dim s As Single, i As Long
    i = PopupSlot(frm)
    If m_closing(i) Then Exit Sub
    If Not m_lost(i) And MenuIsActive(frm) Then frm.TimerInterval = MENU_TICK_MS: Exit Sub
    s = Timer - m_openedAt(i)
    If s < 0 Then s = s + 86400
    If Not m_lost(i) And s * 1000 < MENU_GRACE_MS Then
        On Error Resume Next
        frm.SetFocus
        frm.Controls("txtKeys").SetFocus
        On Error GoTo 0
        frm.TimerInterval = MENU_TICK_MS
        Exit Sub
    End If
    MenuClose frm
End Sub

' Which pop-up's state frm uses: 1 for dlgPopover, 0 for dlgMenu.
Private Function PopupSlot(frm As Form) As Long
    If frm.Name = "dlgPopover" Then PopupSlot = 1
End Function

' True when frm is Access's active form.
Private Function MenuIsActive(frm As Form) As Boolean
    On Error Resume Next
    MenuIsActive = (Screen.ActiveForm.hWnd = frm.hWnd)
End Function

' Build dlgMenu, the pop-up KitMenu opens. Call it once with the app's dialogs.
Public Sub BuildMenuDialog()
    Dim c As Control, r As Long, w As Long, h As Long, y As Long, nm As Variant
    w = px(MENU_MAX_W)
    h = MENU_ROWS * px(MENU_ROW_H) + 2 * px(MENU_PAD)
    NewForm "", 0, w, h
    With F
        .Caption = "Menu"
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
        .Section(acDetail).OnMouseMove = "=MenuHover([Form],-1)"
    End With
    Ev F, "OnLoad", "OnDeactivate", "OnTimer", "OnClose"
    ' Takes the keys: a 1px box with no border, the only Tab stop.
    Set c = Txt("txtKeys", "", 0, 0, 15, 15)
    c.BorderStyle = 0: c.BackStyle = 0: c.OnGotFocus = "": c.OnLostFocus = ""
    Ev c, "OnKeyDown"
    For r = 0 To MENU_ROWS - 1
        y = px(MENU_PAD) + r * px(MENU_ROW_H)
        Rule "ruleSep" & r, px(MENU_PAD), y, w - 2 * px(MENU_PAD), 0, CLR_BORDER
        Set c = Box("boxRow" & r, px(MENU_PAD), y, w - 2 * px(MENU_PAD), px(MENU_ROW_H), CLR_OUT_HOVER)
        c.Visible = False
        ' The caption spans the row (its left margin clears the icon), so the
        ' mouse is over one of the row's parts anywhere in it.
        Set c = Lbl("lblCap" & r, "", px(MENU_PAD), y, w - 2 * px(MENU_PAD), px(MENU_ROW_H), 9, False, CLR_FG)
        c.LeftMargin = px(40)
        VCenter c, px(MENU_ROW_H), 9, True
        Set c = Lbl("lblIcon" & r, "", px(MENU_PAD + 10), y, px(20), px(MENU_ROW_H), 10, False, CLR_MUTED_FG)
        c.FontName = "Segoe MDL2 Assets": c.TextAlign = 2
        VCenter c, px(MENU_ROW_H), 10, True
        For Each nm In Array("boxRow", "lblIcon", "lblCap")
            F.Controls(nm & r).OnClick = "=MenuPick([Form]," & r & ")"
            F.Controls(nm & r).OnMouseMove = "=MenuHover([Form]," & r & ")"
        Next
    Next
    Rule "ruleTop", 0, 15, w, 0, CLR_INPUT_BORDER
    Rule "ruleBottom", 0, h, w, 0, CLR_INPUT_BORDER
    Rule "ruleLeft", 15, 0, 0, h, CLR_INPUT_BORDER
    Rule "ruleRight", w, 0, 0, h, CLR_INPUT_BORDER
    SaveForm "dlgMenu"
End Sub

Private Function ItemText(ByVal i As Long) As String
    ItemText = CStr(m_items(LBound(m_items) + i))
End Function

' An item's caption without its "!" (destructive) mark.
Private Function ItemCaption(ByVal i As Long) As String
    Dim p() As String
    p = Split(ItemText(i), "|")
    ItemCaption = p(1)
    If Left$(ItemCaption, 1) = "!" Then ItemCaption = Mid$(ItemCaption, 2)
End Function

Private Function Pickable(ByVal r As Long) As Boolean
    If r < 0 Or r > UBound(m_items) - LBound(m_items) Then Exit Function
    If ItemText(r) = "-" Then Exit Function
    Pickable = (Left$(ItemText(r), 1) <> "~")
End Function

Private Sub RowColors(frm As Form, ByVal r As Long)
    Dim fore As Long, p() As String
    p = Split(ItemText(r), "|")
    If Left$(p(0), 1) = "~" Then
        fore = Th(CLR_MUTED_FG)
    ElseIf Left$(p(1), 1) = "!" Then
        fore = Th(CLR_DESTRUCTIVE)
    Else
        fore = Th(CLR_FG)
    End If
    frm.Controls("lblCap" & r).ForeColor = fore
    frm.Controls("lblIcon" & r).ForeColor = IIf(fore = Th(CLR_FG), Th(CLR_MUTED_FG), fore)
End Sub

Private Sub RowParts(frm As Form, ByVal r As Long, ByVal v As Boolean)
    frm.Controls("lblIcon" & r).Visible = v
    frm.Controls("lblCap" & r).Visible = v
    If Not v Then frm.Controls("boxRow" & r).Visible = False
End Sub

' The button's place on screen, in pixels.
Private Sub AnchorRect(c As Control)
    Dim frm As Object, pt As POINTAPI, dpi As Long, hdc As LongPtr, top As Long, x As Long, w As Long
    Set frm = c.Parent
    Do While Not TypeOf frm Is Form
        Set frm = frm.Parent
    Loop
    top = c.Top
    If c.Section = acDetail Then top = top + DetailTop(frm)
    hdc = GetDC(0): dpi = GetDeviceCaps(hdc, LOGPIXELSY): ReleaseDC 0, hdc
    If dpi <= 0 Then dpi = 96
    m_hAnchor = frm.hWnd
    ClientToScreen m_hAnchor, pt
    AnchoredX c, frm, x, w
    m_anchor.Left = pt.x + CLng(x * dpi / 1440)
    m_anchor.Top = pt.y + CLng(top * dpi / 1440)
    m_anchor.Right = pt.x + CLng((x + w) * dpi / 1440)
    m_anchor.Bottom = pt.y + CLng((top + c.Height) * dpi / 1440)
End Sub

' Where control c is across frm's window (twips): its Left and Width still
' read their design values when the window is wider than the form, so add
' the growth to a right-anchored control's Left, or to a stretched one's
' Width.
Private Sub AnchoredX(c As Control, frm As Object, x As Long, w As Long)
    Dim grow As Long, ha As Long
    x = c.Left: w = c.Width
    On Error Resume Next
    ha = c.HorizontalAnchor
    grow = frm.InsideWidth - frm.Width
    On Error GoTo 0
    If grow <= 0 Then Exit Sub
    Select Case ha
        Case acHorizontalAnchorRight: x = x + grow
        Case acHorizontalAnchorBoth: w = w + grow
    End Select
End Sub

' Where the current record's Detail section starts in frm's window (twips):
' on a continuous form, that record's row (CurrentSectionTop, which counts
' the header and the scroll), so a button on a card or row anchors to its own
' row; else under the form header.
Private Function DetailTop(frm As Object) As Long
    If frm.DefaultView = 1 Then
        On Error Resume Next
        DetailTop = -1
        DetailTop = frm.CurrentSectionTop
        On Error GoTo 0
        If DetailTop >= 0 Then Exit Function
    End If
    DetailTop = HeaderHeight(frm)
End Function

Private Function HeaderHeight(frm As Object) As Long
    On Error Resume Next
    If frm.Section(acHeader).Visible Then HeaderHeight = frm.Section(acHeader).Height
End Function

' Put frm under the anchor (AnchorRect). show: also show it, on top, without
' taking the focus (a list opened hidden, so typing carries on).
Private Sub Place(frm As Form, Optional ByVal show As Boolean = False)
    Dim rc As RECT, pr As RECT, mi As MONITORINFO, x As Long, y As Long, gap As Long
    GetWindowRect frm.hWnd, rc
    mi.cbSize = LenB(mi)
    If m_hAnchor = 0 Then
        pr = m_anchor: pr.Right = pr.Left + 1: pr.Bottom = pr.Top + 1   ' the pointer's pixel
        If GetMonitorInfo(MonitorFromRect(pr, MONITOR_DEFAULTTONEAREST), mi) = 0 Then Exit Sub
    ElseIf GetMonitorInfo(MonitorFromWindow(m_hAnchor, MONITOR_DEFAULTTONEAREST), mi) = 0 Then
        Exit Sub
    End If
    gap = CLng(MENU_GAP * (rc.Right - rc.Left) / (frm.InsideWidth / 15))
    MenuSpot m_anchor.Left, m_anchor.Top, m_anchor.Right, m_anchor.Bottom, rc.Right - rc.Left, rc.Bottom - rc.Top, _
        mi.rcWork.Left, mi.rcWork.Top, mi.rcWork.Right, mi.rcWork.Bottom, gap, x, y
    If show Then
        SetWindowLongPtr frm.hWnd, GWL_EXSTYLE, GetWindowLongPtr(frm.hWnd, GWL_EXSTYLE) Or WS_EX_NOACTIVATE
        SetWindowPos frm.hWnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Else
        SetWindowPos frm.hWnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
    End If
End Sub

' ---------------------------------------------------------------------------
' Popover: a short explanation next to the control it is about (a web
' popover or rich tooltip), for a "?" or info icon, or a field that needs a
' sentence of help. KitPopover shows it under the control (or above it,
' kept on the monitor) and returns at once; it closes when the user clicks
' anywhere else or presses Esc. dlgPopover (BuildPopoverDialog) is 280px wide
' with a 10pt bold title and the 9pt text, wrapped, up to 8 lines.
' ---------------------------------------------------------------------------


' Show title (or "" for none) and text in a popover by control anchor.
' Returns False, logging why, when dlgPopover isn't built.
Public Function KitPopover(anchor As Control, ByVal title As String, ByVal text As String) As Boolean
    m_popTitle = title
    m_popText = text
    AnchorRect anchor
    If IsFormOpen("dlgPopover") Then DoCmd.Close acForm, "dlgPopover"
    ' After closing the old one, so its Deactivate can't mark this one lost.
    m_closing(1) = False
    m_lost(1) = False
    ' A TitleTip "?" has the same text, so it is covered too.
    HideTip 1, anchor
    On Error GoTo NotBuilt
    DoCmd.OpenForm "dlgPopover"
    KitPopover = True
    Exit Function
NotBuilt:
    RestoreTip 1
    LogError "modMenu.KitPopover", "dlgPopover could not open: " & Err.Description
End Function

' dlgPopover's Close: give the anchor its tooltip back.
Public Sub PopoverClosed()
    RestoreTip 1
End Sub

' dlgMenu's Close: the same for the menu's anchor.
Public Sub MenuClosed()
    RestoreTip 0
End Sub

' Clear anchor's tooltip while pop-up slot shows, noting it for RestoreTip.
Private Sub HideTip(ByVal slot As Long, anchor As Control)
    m_tip(slot) = ""
    On Error Resume Next
    m_tip(slot) = anchor.ControlTipText
    If Len(m_tip(slot)) > 0 Then anchor.ControlTipText = "": Set m_tipAnchor(slot) = anchor
End Sub

' Give the anchor its tooltip back. Its form may have closed first, so a
' failure is ignored.
Private Sub RestoreTip(ByVal slot As Long)
    If m_tipAnchor(slot) Is Nothing Then Exit Sub
    On Error Resume Next
    m_tipAnchor(slot).ControlTipText = m_tip(slot)
    On Error GoTo 0
    Set m_tipAnchor(slot) = Nothing
End Sub

' dlgPopover's Load: fill it, size it to its text and place it.
Public Sub PopoverShow(frm As Form)
    Dim lines As Long, para As Variant, textW As Long, y As Long, h As Long
    textW = px(POPOVER_W) - 2 * px(14)
    For Each para In Split(Replace(m_popText, vbCrLf, vbLf), vbLf)
        If Len(para) = 0 Then
            lines = lines + 1
        Else
            lines = lines - Int(-(MeasureText(CStr(para), 9, 400) * 1.1) / textW)
        End If
    Next
    If lines < 1 Then lines = 1
    If lines > POPOVER_LINES Then lines = POPOVER_LINES
    y = px(12)
    frm.Controls("lblTitle").Visible = (Len(m_popTitle) > 0)
    frm.Controls("lblTitle").Caption = Replace(m_popTitle, "&", "&&")
    If Len(m_popTitle) > 0 Then y = y + px(24)
    h = lines * POPOVER_LINE_H
    With frm.Controls("lblText")
        .Caption = Replace(m_popText, "&", "&&")
        .Top = y
        .Height = h
    End With
    y = y + h + px(12)
    frm.Controls("ruleBottom").Top = y
    frm.Controls("ruleLeft").Height = y
    frm.Controls("ruleRight").Height = y
    frm.InsideWidth = px(POPOVER_W)
    frm.InsideHeight = y
    Place frm
End Sub

' dlgPopover's KeyDown: Esc, Enter and Tab close it.
Public Sub PopoverKey(frm As Form, KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyEscape, vbKeyReturn, vbKeyTab
            KeyCode = 0
            MenuClose frm
    End Select
End Sub

' Build dlgPopover, the pop-up KitPopover opens. Call it once with the app's dialogs.
Public Sub BuildPopoverDialog()
    Dim c As Control, w As Long, h As Long
    w = px(POPOVER_W)
    h = px(12) + px(24) + POPOVER_LINES * POPOVER_LINE_H + px(12)
    NewForm "", 0, w, h
    With F
        .Caption = "Popover"
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
        .Section(acDetail).OnClick = "=MenuClose([Form])"
    End With
    Ev F, "OnLoad", "OnDeactivate", "OnTimer", "OnClose"
    Set c = Txt("txtKeys", "", 0, 0, 15, 15)
    c.BorderStyle = 0: c.BackStyle = 0: c.OnGotFocus = "": c.OnLostFocus = ""
    Ev c, "OnKeyDown"
    Set c = Lbl("lblTitle", "", px(14), px(12), w - 2 * px(14), px(20), 10, True, CLR_FG)
    c.OnClick = "=MenuClose([Form])"
    Set c = Lbl("lblText", "", px(14), px(36), w - 2 * px(14), POPOVER_LINES * POPOVER_LINE_H, 9, False, CLR_FG)
    c.TopMargin = 0
    c.OnClick = "=MenuClose([Form])"
    Rule "ruleTop", 0, 15, w, 0, CLR_INPUT_BORDER
    Rule "ruleBottom", 0, h, w, 0, CLR_INPUT_BORDER
    Rule "ruleLeft", 15, 0, 0, h, CLR_INPUT_BORDER
    Rule "ruleRight", w, 0, 0, h, CLR_INPUT_BORDER
    SaveForm "dlgPopover"
End Sub

' ---------------------------------------------------------------------------
' Searchable picker: a text box the user types a few letters into, with the
' matching rows dropping down under it (a web combobox or autocomplete), for
' a list too long to scroll: people, hull numbers, templates. PickBox builds
' it. The list is dlgPick (BuildPickDialog), a borderless pop-up under the
' box with the first PICK_ROWS matches; it never takes the focus, so the
' user keeps typing. The top match is lit: Up and Down move the light, Enter
' picks the lit row, a click picks a row, Esc closes the list, and Down on
' a closed list opens it. Leaving the box closes the list too: the box goes
' back to the picked row's text, and a box left empty clears the pick.
'
' rowSql gives the rows: its first column is the key, its second is what the
' box shows once a row is picked, and every column after the key is
' searched with LIKE *text*. The list shows the second column, with any
' further columns in muted text on the right. Rows come ordered by the
' second column, so leave ORDER BY out of rowSql. The key goes in a hidden
' box beside it, <name>Key, bound to keyField ("" for an unbound dialog):
' read the pick there. The form's code adds (and a form the kit's builders
' make sets Ev c, "OnKeyDown" on the box PickBox returns):
'   Private Sub txtPerson_KeyDown(KeyCode As Integer, Shift As Integer)
'       KitPickKey Me.txtPerson, KeyCode
'   End Sub
' and KitPickShow Me.txtPerson in Form_Current on a bound form (an unbound
' dialog calls it after setting txtPersonKey), so the box shows the pick.
' ---------------------------------------------------------------------------

' Build a searchable picker at x, y, w wide (h tall, 30px by default): the
' box nm the user types in, with placeholder showing while it is empty,
' and the hidden nm & "Key" bound to keyField ("" for none), which holds
' the pick. See the top of this section for rowSql and the form's code.
' Raises error 5 when rowSql has an ORDER BY.
Public Function PickBox(ByVal nm As String, ByVal keyField As String, ByVal rowSql As String, ByVal x As Long, _
        ByVal y As Long, ByVal w As Long, Optional ByVal h As Long = 450, _
        Optional ByVal placeholder As String = "Type to search") As Control
    Dim c As Control
    If InStr(1, rowSql, "ORDER BY", vbTextCompare) > 0 Then _
        Err.Raise 5, , "PickBox " & nm & ": leave ORDER BY out of rowSql (the list is ordered by its second column)"
    Set c = Txt(nm & "Key", keyField, x, y, w, h)
    c.Visible = False: c.TabStop = False
    Set c = Txt(nm, "", x, y, w, h)
    c.Format = "@;""" & Replace(placeholder, """", "") & """"
    c.Tag = PICK_TAG & rowSql
    c.AllowAutoCorrect = False
    c.OnChange = "=KitPickChanged([Form],""" & nm & """)"
    Set PickBox = c
End Function

' What the user has typed in box. An empty box's Text while it has the focus
' is its placeholder (the Format's second section), which is not typing.
Private Function PickTyped(box As Control) As String
    Dim t As String, f As String, i As Long
    t = Nz(box.Text, "")
    f = Nz(box.Format, "")
    i = InStr(f, ";")
    If i > 0 And Len(t) > 0 Then
        If t = Replace(Mid$(f, i + 1), """", "") Then t = ""
    End If
    PickTyped = t
End Function

' A PickBox's Change (PickBox sets it): list the rows matching what is typed.
Public Function KitPickChanged(frm As Form, ByVal nm As String)
    On Error GoTo Fail
    PickList frm.Controls(nm), PickTyped(frm.Controls(nm))
    Exit Function
Fail:
    If PickCanceled(Err.Number) Then PickEscaped: Exit Function
    HandleError "modMenu.KitPickChanged", nm
End Function

' True for error n when it is the user's Esc, not a fault: Access cancels a
' query that is running when Esc is pressed (3059), and the list's query runs
' on each key typed. Public so the kit tests can check it.
Public Function PickCanceled(ByVal n As Long) As Boolean
    PickCanceled = (n = 3059)
End Function

' The user's Esc cancelled the list's query (PickCanceled): close the list
' as Esc does. Putting the box back runs a query too, so the same Esc may
' cancel that as well; the box then keeps its text.
' Its KeyDown may still follow; KitPickKey takes it then, so the box doesn't
' undo the typing and open the list again.
Private Sub PickEscaped()
    m_pickEscAt = Timer
    On Error Resume Next
    PickClose True
End Sub

' A PickBox's KeyDown (see the top of this section): Up and Down move the
' lit row (Down opens a closed list), Enter picks the lit row, Esc closes
' the list. Other keys are left to the box.
Public Sub KitPickKey(box As Control, KeyCode As Integer)
    On Error GoTo Fail
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            If Not PickOpenFor(box) Then
                If KeyCode = vbKeyDown Then PickList box, PickTyped(box): KeyCode = 0
                Exit Sub
            End If
            If m_pickKeys.Count > 0 Then
                If KeyCode = vbKeyDown Then
                    If m_pickHot < m_pickKeys.Count - 1 Then PickLight Forms("dlgPick"), m_pickHot + 1
                ElseIf m_pickHot > 0 Then
                    PickLight Forms("dlgPick"), m_pickHot - 1
                End If
            End If
            KeyCode = 0
        Case vbKeyReturn
            If PickOpenFor(box) Then
                If m_pickHot >= 0 Then PickChoose m_pickHot
                KeyCode = 0
            End If
        Case vbKeyEscape
            If PickOpenFor(box) Then
                PickClose True: KeyCode = 0
            ElseIf Abs(Timer - m_pickEscAt) < 0.5 Then
                KeyCode = 0         ' the Esc PickEscaped has closed the list for
            End If
    End Select
    Exit Sub
Fail:
    If PickCanceled(Err.Number) Then PickEscaped: KeyCode = 0: Exit Sub
    HandleError "modMenu.KitPickKey", box.Name
End Sub

' Show the picked row's text in box: the second column of rowSql's row for
' the key in its <name>Key, or nothing when there is no pick (or the key's
' row is gone).
Public Sub KitPickShow(box As Control)
    Dim k As Variant, db As DAO.Database, qd As DAO.QueryDef, rs As DAO.Recordset, cols As Variant
    On Error GoTo Fail
    k = PickForm(box).Controls(box.Name & "Key").Value
    If IsNull(k) Then box.Value = Null: Exit Sub
    cols = PickColumns(box)
    Set db = CurrentDb                  ' held: a query def needs its database open
    Set qd = db.CreateQueryDef("", "SELECT q.* FROM (" & PickSource(box) & ") AS q WHERE q.[" & cols(0) & "]=[pickKey]")
    qd.Parameters("pickKey") = k
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    If rs.EOF Then box.Value = Null Else box.Value = rs.Fields(1).Value
    rs.Close
    Exit Sub
Fail:
    RaiseAgain "modMenu.KitPickShow", box.Name
End Sub

' text for a LIKE "*...*" pattern: Access's wildcards ([ * ? #) taken
' literally. Public so the kit tests can check it.
Public Function PickLikeText(ByVal text As String) As String
    text = Replace(text, "[", "[[]")
    text = Replace(text, "*", "[*]")
    text = Replace(text, "?", "[?]")
    PickLikeText = Replace(text, "#", "[#]")
End Function

' The list's query for box: the first PICK_ROWS rows of its rowSql where
' any column after the key is LIKE the parameter [pickText] (all rows when
' nothing is typed), by the second column, then the key. Public so the kit
' tests can check it.
Public Function PickSql(box As Control, ByVal typed As Boolean) As String
    Dim cols As Variant, i As Long, where As String
    cols = PickColumns(box)
    If typed Then
        For i = 1 To UBound(cols)
            where = where & IIf(Len(where) > 0, " OR ", " WHERE ") & "q.[" & cols(i) & "] LIKE [pickText]"
        Next
    End If
    PickSql = "SELECT TOP " & PICK_ROWS & " q.* FROM (" & PickSource(box) & ") AS q" & _
        where & " ORDER BY q.[" & cols(1) & "], q.[" & cols(0) & "]"
End Function

' dlgPick's row click (set by BuildPickDialog): pick row r.
Public Function PickClick(frm As Form, ByVal r As Long)
    On Error GoTo Fail
    If m_pickKeys Is Nothing Then Exit Function
    If r >= 0 And r < m_pickKeys.Count Then PickChoose r
    Exit Function
Fail:
    HandleError "modMenu.PickClick", "row " & r
End Function

' dlgPick's mouse move (set by BuildPickDialog): light row r (-1 for none
' keeps the light where it is, so Enter still picks it).
Public Function PickHover(frm As Form, ByVal r As Long)
    On Error GoTo Fail
    If m_pickKeys Is Nothing Then Exit Function
    If r >= 0 And r < m_pickKeys.Count Then PickLight frm, r
    Exit Function
Fail:
    LogError "modMenu.PickHover", "row " & r
End Function

' dlgPick's timer (set by BuildPickDialog): once the box has lost the focus
' (the user clicked or tabbed elsewhere, or VBA was reset), close the list
' and put the box back to its pick.
Public Function PickTick(frm As Form)
    On Error GoTo Fail
    If PickBoxFocused() Then Exit Function
    PickClose True, True
    Exit Function
Fail:
    LogError "modMenu.PickTick"
    On Error Resume Next
    DoCmd.Close acForm, frm.Name
End Function

' Build dlgPick, the list a PickBox drops down. Call it once with the app's dialogs.
Public Sub BuildPickDialog()
    Dim c As Control, r As Long, w As Long, h As Long, y As Long, nm As Variant
    w = px(PICK_MAX_W)
    h = PICK_ROWS * px(PICK_ROW_H) + 2 * px(MENU_PAD)
    NewForm "", 0, w, h
    With F
        .Caption = "Pick"
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
        .OnTimer = "=PickTick([Form])"
        .Section(acDetail).BackColor = CLR_DIALOG
        .Section(acDetail).AlternateBackColor = CLR_DIALOG
    End With
    For r = 0 To PICK_ROWS - 1
        y = px(MENU_PAD) + r * px(PICK_ROW_H)
        Set c = Box("boxRow" & r, px(MENU_PAD), y, w - 2 * px(MENU_PAD), px(PICK_ROW_H), CLR_OUT_HOVER)
        c.Visible = False
        ' The main text spans the row, so the mouse is over one of the
        ' row's parts anywhere in it; the muted text sits over its right end.
        Set c = Lbl("lblMain" & r, "", px(MENU_PAD), y, w - 2 * px(MENU_PAD), px(PICK_ROW_H), 9, False, CLR_FG)
        c.LeftMargin = px(12): c.RightMargin = px(12)
        VCenter c, px(PICK_ROW_H), 9, True
        Set c = Lbl("lblSub" & r, "", w - px(MENU_PAD) - px(160), y, px(160), px(PICK_ROW_H), 9, False, CLR_MUTED_FG)
        c.TextAlign = 3: c.RightMargin = px(12)
        VCenter c, px(PICK_ROW_H), 9, True
        For Each nm In Array("boxRow", "lblMain", "lblSub")
            F.Controls(nm & r).OnClick = "=PickClick([Form]," & r & ")"
            F.Controls(nm & r).OnMouseMove = "=PickHover([Form]," & r & ")"
        Next
    Next
    Rule "ruleTop", 0, 15, w, 0, CLR_INPUT_BORDER
    Rule "ruleBottom", 0, h, w, 0, CLR_INPUT_BORDER
    Rule "ruleLeft", 15, 0, 0, h, CLR_INPUT_BORDER
    Rule "ruleRight", w, 0, 0, h, CLR_INPUT_BORDER
    SaveForm "dlgPick"
End Sub

' List the rows of box's rowSql matching text under box, opening dlgPick
' (hidden, then shown without taking the focus) when it isn't open.
Private Sub PickList(box As Control, ByVal text As String)
    Dim db As DAO.Database, qd As DAO.QueryDef, rs As DAO.Recordset, frm As Form, n As Long, i As Long, s As String
    Dim subs(0 To PICK_ROWS - 1) As String
    If Not m_pickBox Is Nothing Then
        If Not SameBox(m_pickBox, box) Then PickClose True, True    ' another box's list
    End If
    text = Trim$(text)
    Set db = CurrentDb                  ' held: a query def needs its database open
    Set qd = db.CreateQueryDef("", PickSql(box, Len(text) > 0))
    If Len(text) > 0 Then qd.Parameters("pickText") = "*" & PickLikeText(text) & "*"
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    If Not IsFormOpen("dlgPick") Then DoCmd.OpenForm "dlgPick", WindowMode:=acHidden
    Set frm = Forms("dlgPick")
    Set m_pickBox = box
    Set m_pickKeys = New Collection
    Set m_pickText = New Collection
    Do While Not rs.EOF And n < PICK_ROWS
        m_pickKeys.Add rs.Fields(0).Value
        m_pickText.Add Nz(rs.Fields(1).Value, "")
        s = ""
        For i = 2 To rs.Fields.Count - 1
            If Len(Nz(rs.Fields(i).Value, "")) > 0 Then s = s & IIf(Len(s) > 0, "  " & ChrW(&HB7) & "  ", "") & rs.Fields(i).Value
        Next
        subs(n) = s
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    PickLayout frm, box, n, subs
End Sub

' Size dlgPick to n rows (one "No matches" row for none) as wide as box,
' within PICK_MIN_W and PICK_MAX_W, light the first row, and show it under
' box (above it when there's no room below) without taking the focus.
Private Sub PickLayout(frm As Form, box As Control, ByVal n As Long, subs() As String)
    Dim r As Long, w As Long, y As Long, subW As Long, shown As Long, mainW As Long
    w = box.Width
    If w < px(PICK_MIN_W) Then w = px(PICK_MIN_W)
    If w > px(PICK_MAX_W) Then w = px(PICK_MAX_W)
    shown = IIf(n = 0, 1, n)
    If n = 0 Then subs(0) = ""
    frm.Controls("lblMain0").ForeColor = IIf(n = 0, Th(CLR_MUTED_FG), Th(CLR_FG))
    For r = 0 To PICK_ROWS - 1
        frm.Controls("boxRow" & r).Visible = False
        frm.Controls("lblMain" & r).Visible = (r < shown)
        frm.Controls("lblSub" & r).Visible = (r < shown)
        If r < shown Then
            ' The muted text takes what it needs, up to 40% of the row.
            subW = 0
            If Len(subs(r)) > 0 Then subW = MeasureText(subs(r), 9, 400) + px(24)
            If subW > (w - 2 * px(MENU_PAD)) * 2 \ 5 Then subW = (w - 2 * px(MENU_PAD)) * 2 \ 5
            ' One line each: text too long for its part ends in an ellipsis
            ' (a label would wrap it onto a second line the row cuts off).
            mainW = w - 2 * px(MENU_PAD) - px(24) - subW
            If n = 0 Then
                frm.Controls("lblMain0").Caption = "No matches"
            Else
                frm.Controls("lblMain" & r).Caption = Replace(FitSingleLineText(m_pickText(r + 1), mainW), "&", "&&")
            End If
            frm.Controls("lblSub" & r).Caption = Replace(FitSingleLineText(subs(r), subW - px(24)), "&", "&&")
            frm.Controls("boxRow" & r).Width = w - 2 * px(MENU_PAD)
            frm.Controls("lblMain" & r).Width = w - 2 * px(MENU_PAD)
            frm.Controls("lblMain" & r).RightMargin = px(12) + subW
            frm.Controls("lblSub" & r).Width = IIf(subW > 0, subW, px(24))
            frm.Controls("lblSub" & r).Left = w - px(MENU_PAD) - frm.Controls("lblSub" & r).Width
        End If
    Next
    y = px(MENU_PAD) + shown * px(PICK_ROW_H) + px(MENU_PAD)
    frm.Controls("ruleTop").Width = w
    frm.Controls("ruleBottom").Top = y: frm.Controls("ruleBottom").Width = w
    frm.Controls("ruleLeft").Height = y
    frm.Controls("ruleRight").Left = w: frm.Controls("ruleRight").Height = y
    frm.InsideWidth = w
    frm.InsideHeight = y
    m_pickHot = -1
    If n > 0 Then PickLight frm, 0
    AnchorRect box
    Place frm, True
    frm.TimerInterval = PICK_TICK_MS
End Sub

' Light row r of the list.
Private Sub PickLight(frm As Form, ByVal r As Long)
    If r = m_pickHot Then Exit Sub
    If m_pickHot >= 0 Then frm.Controls("boxRow" & m_pickHot).Visible = False
    m_pickHot = r
    If r >= 0 Then frm.Controls("boxRow" & r).Visible = True
End Sub

' Pick row r: its key into the box's <name>Key, its text into the box, and
' close the list. The focus goes back to the box (a click on the list may
' have taken it), with the cursor at the end.
Private Sub PickChoose(ByVal r As Long)
    Dim box As Control
    Set box = m_pickBox
    PickForm(box).Controls(box.Name & "Key").Value = m_pickKeys(r + 1)
    box.Value = m_pickText(r + 1)
    PickClose False
    On Error Resume Next                ' a box that can't take the focus keeps the pick all the same
    box.SetFocus
    box.SelStart = Len(PickTyped(box))
End Sub

' Close the list. restore: put the box back to its pick (Esc, or leaving
' the box without picking). clearEmpty: a box left empty clears the pick
' (leaving it; Esc keeps the pick).
Private Sub PickClose(ByVal restore As Boolean, Optional ByVal clearEmpty As Boolean = False)
    Dim box As Control, keyBox As Control, typed As String
    Set box = m_pickBox
    Set m_pickBox = Nothing
    m_pickHot = -1
    If IsFormOpen("dlgPick") Then
        Forms("dlgPick").TimerInterval = 0
        DoCmd.Close acForm, "dlgPick"
    End If
    If Not restore Or box Is Nothing Then Exit Sub
    On Error GoTo Gone                  ' the box's form closed while the list was open
    Set keyBox = PickForm(box).Controls(box.Name & "Key")
    If PickBoxHasFocus(box) Then typed = PickTyped(box) Else typed = Nz(box.Value, "")
    If clearEmpty And Len(Trim$(typed)) = 0 And Not IsNull(keyBox.Value) Then keyBox.Value = Null
    KitPickShow box
Gone:
End Sub

' True when the open list belongs to box.
Private Function PickOpenFor(box As Control) As Boolean
    If m_pickBox Is Nothing Then Exit Function
    If Not IsFormOpen("dlgPick") Then Exit Function
    On Error GoTo Gone
    PickOpenFor = SameBox(m_pickBox, box)
Gone:
End Function

' True when a and b are the same control on the same open form.
Private Function SameBox(a As Control, b As Control) As Boolean
    On Error GoTo Gone                  ' a's form has closed
    SameBox = (a.Name = b.Name And PickForm(a).hWnd = PickForm(b).hWnd)
Gone:
End Function

' True while the open list's box has the focus, or the list itself does
' (a click on a row is on its way).
Private Function PickBoxFocused() As Boolean
    If m_pickBox Is Nothing Then Exit Function
    On Error GoTo Nope
    If Screen.ActiveForm.Name = "dlgPick" Then PickBoxFocused = True: Exit Function
    PickBoxFocused = PickBoxHasFocus(m_pickBox)
Nope:
End Function

Private Function PickBoxHasFocus(box As Control) As Boolean
    Dim a As Control
    On Error GoTo Nope
    Set a = Screen.ActiveControl
    PickBoxHasFocus = SameBox(a, box)
Nope:
End Function

' The form a control is on (its Parent is a page when it sits on a tab).
Private Function PickForm(c As Control) As Form
    Dim p As Object
    Set p = c.Parent
    Do While Not TypeOf p Is Form
        Set p = p.Parent
    Loop
    Set PickForm = p
End Function

' The rowSql a PickBox keeps in its Tag. Error 5 for a box PickBox didn't build.
Private Function PickSource(box As Control) As String
    If Left$(box.Tag, Len(PICK_TAG)) <> PICK_TAG Then Err.Raise 5, , box.Name & " is not a PickBox"
    PickSource = Mid$(box.Tag, Len(PICK_TAG) + 1)
End Function

' rowSql's column names, the key first. Error 5 for fewer than 2 columns.
Private Function PickColumns(box As Control) As Variant
    Dim rs As DAO.Recordset, i As Long, out() As String
    Set rs = CurrentDb.OpenRecordset("SELECT q.* FROM (" & PickSource(box) & ") AS q WHERE False", dbOpenSnapshot)
    If rs.Fields.Count < 2 Then rs.Close: Err.Raise 5, , box.Name & ": rowSql needs a key and a column to show"
    ReDim out(0 To rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        out(i) = rs.Fields(i).Name
    Next
    rs.Close
    PickColumns = out
End Function
