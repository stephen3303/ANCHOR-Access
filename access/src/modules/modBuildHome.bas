Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Home page (the web app's /): four tiles
' (unread, acknowledgements due, action items open, status changes), the
' messages that require the person's acknowledgement, the latest traffic,
' and the status changes on the right. Built by modBuild.BuildPages
' (BuildHomePage builds the tables first, then the page).
'
' Owner: the Home page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW); the kit's modUIKit.

Private Const HOME_SIDE_W As Long = 5400        ' the Status changes card, 360 px

Private Function HomeMainW() As Long
    HomeMainW = ContentW() - HOME_SIDE_W - px(12)
End Function

Public Sub BuildHomePage()
    BuildHomeGrids
    Build_pgHome
End Sub

' Every table Home shows, before the form that hosts them.
Public Sub BuildHomeGrids()
    BuildGrid "subGridHomeAck", HomeAckCols(), HomeMainW() - px(32)
    BuildGrid "subGridHomeLatest", HomeLatestCols(), HomeMainW() - px(32)
    BuildGrid "subGridHomeStatus", HomeStatusCols(), HOME_SIDE_W - px(32), 2
End Sub

' "field|caption|width px|options[|sort field]"
Private Function HomeAckCols() As Variant
    HomeAckCols = Array("subject|Subject|236|b", "ref|Reference|136|", "prec|Precedence|96|p|prec_rank", "due_txt|Due|120|m|due")
End Function

Private Function HomeLatestCols() As Variant
    HomeLatestCols = Array("subject|Subject|236|", "ref|Reference|136|", "dtg|DTG|120|m|received", "prec|Precedence|96|p|prec_rank")
End Function

Private Function HomeStatusCols() As Variant
    HomeStatusCols = Array("lifecycle|Status|104|p", "item_txt|Message|224|")
End Function

' The tiles over the acknowledgements and the latest traffic on the left
' and the status changes on the right; a click on a row opens its message.
' The tiles' values and the welcome line are set in code.
Public Sub Build_pgHome()
    Dim c As Control, x As Long, y As Long, h As Long, tw As Long, i As Long, mw As Long
    Dim tiles As Variant, labels As Variant, glyphs As Variant
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnTimer"             ' OnTimer: the tiles and tables load once the page is open
    PageHeader "Home", "What needs you now: the messages you must acknowledge, the messages whose status changed " & _
        "and the latest traffic. Click a message to read it.", PAGE_PAD
    tiles = Array("tileUnread", "tileAck", "tileAction", "tileStatus")
    labels = Array("Unread", "Acknowledgements due", "Action items open", "Status changes")
    glyphs = Array("E715", "EA8F", "E9D5", "E7BA")
    tw = (ContentW() - 3 * px(12)) \ 4
    For i = 0 To 3
        x = PAGE_PAD + i * (tw + px(12))
        StatTile CStr(tiles(i)), CStr(labels(i)), "", x, px(84), IIf(i = 3, ContentW() - 3 * (tw + px(12)), tw), px(112), _
            caption:=" ", glyph:=CStr(glyphs(i))
    Next

    y = px(208): mw = HomeMainW(): h = px(196)
    Card "Ack", "Requires your acknowledgement", PAGE_PAD, y, mw, h, , 2, 0
    Set c = Lbl("lblAckDone", "", PAGE_PAD + mw - px(16) - px(200), y + px(10), px(200), px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 3: Anchor c, 1, 0
    GridHeaders "hdrAck", HomeAckCols(), PAGE_PAD + px(16), y + px(40), w:=mw - px(32), sortSf:="sfAck"
    Set c = SubF("sfAck", "subGridHomeAck", PAGE_PAD + px(16), y + px(62), mw - px(32), h - px(78)): Anchor c, 2, 0
    EmptyState "emptyAck", PAGE_PAD + px(16), y + px(40), mw - px(32), "E930", "Nothing to acknowledge", _
        "Designated traffic you must acknowledge shows here until you do."

    y = y + h + px(12): h = PAGE_H - px(16) - y
    Card "Latest", "Latest traffic", PAGE_PAD, y, mw, h, , 2, 2
    Set c = Link("lnkAll", "View all", PAGE_PAD + mw - px(16) - px(80), y + px(10), px(80), px(20), "=NavTo(""messages"")", 9, CLR_LINK)
    c.TextAlign = 3: Anchor c, 1, 0
    GridHeaders "hdrLatest", HomeLatestCols(), PAGE_PAD + px(16), y + px(40), w:=mw - px(32), sortSf:="sfLatest"
    Set c = SubF("sfLatest", "subGridHomeLatest", PAGE_PAD + px(16), y + px(62), mw - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyLatest", PAGE_PAD + px(16), y + px(40), mw - px(32), "E715", "No traffic yet", _
        "Messages ANCHOR receives from C2OIX show here, newest first."

    x = PAGE_PAD + mw + px(12): y = px(208): h = PAGE_H - px(16) - y
    Card "Status", "Status changes", x, y, HOME_SIDE_W, h, , 1, 2
    Set c = Lbl("lblStatusSub", "Messages amended, corrected, superseded, cancelled or expired, newest first.", _
        x + px(16), y + px(36), HOME_SIDE_W - px(32), px(32), 8, False, CLR_MUTED_FG)
    c.TopMargin = 0: Anchor c, 1, 0
    GridHeaders "hdrStatus", HomeStatusCols(), x + px(16), y + px(72), w:=HOME_SIDE_W - px(32)
    Set c = SubF("sfStatus", "subGridHomeStatus", x + px(16), y + px(94), HOME_SIDE_W - px(32), h - px(110)): Anchor c, 1, 2
    EmptyState "emptyStatus", x + px(16), y + px(72), HOME_SIDE_W - px(32), "E930", "No status changes", _
        "Every message ANCHOR holds is the current version."
    SaveForm "pgHome"
End Sub
