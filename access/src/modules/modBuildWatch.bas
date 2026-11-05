Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Watch turnover page (/watch-turnover): watch notes, ongoing topics and the turnover checklist.
'
' Four tiles across the top (the web app's watch prep counts: unread
' messages, acknowledgements due, open action items, and the checklist's
' progress); the watch notes and the ongoing topics on the left, the
' turnover checklist with Start and Complete turnover on the right.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW); the kit's modUIKit.

Private Const SIDE_W As Long = 5850             ' the checklist card, 390 px
Private Const CARD_Y As Long = 2640             ' the cards' top, 176 px

Private Function LeftW() As Long
    LeftW = ContentW() - SIDE_W - px(12)
End Function

' The Watch turnover page and the tables it hosts (modBuild.BuildPages).
Public Sub BuildWatchPage()
    BuildWatchGrids
    Build_pgWatchTurnover
End Sub

' The page's tables, before the page.
Public Sub BuildWatchGrids()
    BuildGrid "subGridWatchNotes", WatchNoteCols(), LeftW() - px(32)
    BuildGrid "subGridWatchTopics", WatchTopicCols(), LeftW() - px(32)
    BuildGrid "subGridWatchChecks", WatchCheckCols(), SIDE_W - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function WatchNoteCols() As Variant
    WatchNoteCols = Array("note_state|Status|80|p", "created_txt|Added|96|m|CreatedOn", "who|By|100|m", "note_text|Note|200|")
End Function

Private Function WatchTopicCols() As Variant
    WatchTopicCols = Array("topic_name|Topic|240|b", "topic_status|Status|104|p", "upd_txt|Last update|160|m|LastUpdate")
End Function

Private Function WatchCheckCols() As Variant
    WatchCheckCols = Array("check_state|Status|104|p", "checked_txt|Ticked|56|m", "item_text|Item|160|")
End Function

' The tiles (values set in code); the watch notes card with Add note and
' Close; the ongoing topics card with Add topic and Status; the checklist
' card: the turnover's line, its progress, the items (a click ticks one) and
' Start turnover and Complete turnover.
Public Sub Build_pgWatchTurnover()
    Dim c As Control, x As Long, y As Long, h As Long, tw As Long, i As Long, lw As Long, hh As Long, fx As Long, fw As Long
    Dim tiles As Variant, labels As Variant, glyphs As Variant, k As String
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Watch turnover", "Review traffic. Resolve priorities. Take the watch. The counts are yours; the notes, " & _
        "topics and checklist are the watch's, shared by everyone who stands it.", PAGE_PAD

    tiles = Array("tileUnread", "tileAcks", "tileActions", "tileChecks")
    labels = Array("Unread messages", "Acknowledgements due", "Open action items", "Turnover checklist")
    glyphs = Array("E715", "E73E", "E8FB", "E9D5")
    tw = (ContentW() - 3 * px(12)) \ 4
    For i = 0 To 3
        x = PAGE_PAD + i * (tw + px(12))
        StatTile CStr(tiles(i)), CStr(labels(i)), "", x, px(84), IIf(i = 3, ContentW() - 3 * (tw + px(12)), tw), px(80), _
            glyph:=CStr(glyphs(i)), fmt:=IIf(i = 3, "", "#,##0")
    Next

    y = CARD_Y: h = PAGE_H - y - px(16): lw = LeftW()
    hh = (h - px(12)) \ 2
    ' Watch notes.
    Card "Notes", "Watch notes", PAGE_PAD, y, lw, hh, , 2, 0
    Set c = Btn("btnCloseNote", "Close", PAGE_PAD + lw - px(16) - px(72), y + px(8), px(72), px(28), "ghost"): Anchor c, 1, 0
    c.ControlTipText = "Close the picked note: it is dealt with"
    Set c = Btn("btnAddNote", "Add note", PAGE_PAD + lw - px(16) - px(160), y + px(8), px(80), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Add a note for the watch"
    GridHeaders "hdrNotes", WatchNoteCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfNotes"
    Set c = SubF("sfNotes", "subGridWatchNotes", PAGE_PAD + px(16), y + px(62), lw - px(32), hh - px(78)): Anchor c, 2, 0
    EmptyState "emptyNotes", PAGE_PAD + px(16), y + px(70), lw - px(32), "E70B", "No watch notes", _
        "Notes the watch leaves for the next one show here."
    ' Ongoing topics.
    y = CARD_Y + hh + px(12): hh = h - hh - px(12)
    Card "Topics", "Ongoing topics", PAGE_PAD, y, lw, hh, , 2, 2
    Set c = Btn("btnStatus", "Status", PAGE_PAD + lw - px(16) - px(72), y + px(8), px(72), px(28), "ghost"): Anchor c, 1, 0
    c.ControlTipText = "Set the picked topic Active, Monitoring or Closed"
    Set c = Btn("btnAddTopic", "Add topic", PAGE_PAD + lw - px(16) - px(168), y + px(8), px(88), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Add a topic the watch is following"
    GridHeaders "hdrTopics", WatchTopicCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfTopics"
    Set c = SubF("sfTopics", "subGridWatchTopics", PAGE_PAD + px(16), y + px(62), lw - px(32), hh - px(78)): Anchor c, 2, 2
    EmptyState "emptyTopics", PAGE_PAD + px(16), y + px(70), lw - px(32), "E8F1", "No ongoing topics", _
        "Topics the watch is following show here."

    ' The turnover checklist.
    x = PAGE_PAD + lw + px(12): y = CARD_Y: fx = x + px(16): fw = SIDE_W - px(32)
    Card "Checks", "Turnover checklist", x, y, SIDE_W, h, , 1, 2
    k = BuiltNames()
    Lbl "lblTurn", "", fx, y + px(40), fw, px(36), 9, False, CLR_MUTED_FG          ' two lines: the turnover's state
    Set c = RO("txtPct", "", fx, y + px(40), px(4), px(4)): c.Visible = False      ' the progress bar's percent (set in code)
    Progress "prgChecks", "[txtPct]", fx, y + px(84), fw, px(8)
    GridHeaders "hdrChecks", WatchCheckCols(), fx, y + px(100), w:=fw
    AnchorBuilt k, 1, 0
    Set c = SubF("sfChecks", "subGridWatchChecks", fx, y + px(122), fw, h - px(122) - px(64)): Anchor c, 1, 2
    k = BuiltNames()
    EmptyState "emptyChecks", fx, y + px(130), fw, "E9D5", "No checklist items", _
        "An admin keeps the turnover checklist's items."
    AnchorBuilt k, 1, 0
    Set c = Btn("btnStart", "Start turnover", fx, y + h - px(48), px(136), px(32)): Anchor c, 1, 1
    c.ControlTipText = "Start a turnover, so the checklist can be ticked for it"
    Set c = Btn("btnComplete", "Complete turnover", fx + px(144), y + h - px(48), px(152), px(32)): Anchor c, 1, 1
    c.ControlTipText = "Complete the turnover once every item is ticked"
    SaveForm "pgWatchTurnover"
End Sub
