Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Reports page (/reports): acknowledgement accountability and traffic measures.
'
' The period over four tiles (messages received, the share of
' acknowledgements made, those still owed, and overdue messages); under
' them the accountability table with the picked message's people or
' commands below it, and the traffic measures on the right.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, FieldLbl); the kit's modUIKit.

Private Const SIDE_W As Long = 5400             ' the traffic measures card, 360 px
Private Const CARD_Y As Long = 3300             ' the two cards' top, 220 px

Private Function MainW() As Long
    MainW = ContentW() - SIDE_W - px(12)
End Function

' The Reports page and the tables it hosts (modBuild.BuildPages).
Public Sub BuildReportsPage()
    BuildReportsGrids
    Build_pgReports
End Sub

' The page's tables, before the page.
Public Sub BuildReportsGrids()
    BuildGrid "subGridRptAck", RptAckCols(), MainW() - px(32)
    BuildGrid "subGridRptAckDet", RptDetCols(), MainW() - px(32)
    BuildGrid "subGridRptTraffic", RptTrafficCols(), SIDE_W - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function RptAckCols() As Variant
    RptAckCols = Array("ref|Reference|152|b", "prec|Precedence|96|p|prec_rank", "ack_txt|Acknowledged|112|r|acked", _
        "status|Status|96|p", "subject_txt|Subject|132|")
End Function

Private Function RptDetCols() As Variant
    RptDetCols = Array("who|Person or command|200|b", "via_txt|Unit, or role and via|160|m", "status|Status|104|p", _
        "when_txt|Acknowledged|120|m|ack_on")
End Function

Private Function RptTrafficCols() As Variant
    RptTrafficCols = Array("label|Value|128|p", "n|Count|56|r", "share_txt|Share|56|r", "note_txt|Note|88|m")
End Function

' The period (Segmented: 7 days, 30 days, All) and what it covers; four
' tiles whose values are set in code; the accountability table with Open
' message, the picked message's acknowledgements under it; the traffic
' measures card with its own Segmented (precedence, type, class,
' lifecycle) on the right. Export CSV at the top right writes the
' accountability table.
Public Sub Build_pgReports()
    Dim c As Control, x As Long, y As Long, h As Long, tw As Long, i As Long, mw As Long, k As String
    Dim tiles As Variant, labels As Variant, glyphs As Variant
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Reports", "Who has acknowledged the traffic that must be acknowledged: every ANCHOR user for traffic we " & _
        "received, the addressee commands for traffic our command sent. And how much traffic arrived, by precedence, " & _
        "type, classification and lifecycle.", PAGE_PAD
    Set c = Btn("btnExport", "Export CSV", PAGE_W - PAGE_PAD - px(120), px(24), px(120), px(32), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Save the accountability table as a CSV file"

    y = px(84)
    Segmented "segPeriod", "", PAGE_PAD, y, px(300), RPT_PERIOD_7 & "|7 days;" & RPT_PERIOD_30 & "|30 days;" & _
        RPT_PERIOD_ALL & "|All", px(32), "PeriodChanged"
    Lbl "lblPeriod", "", PAGE_PAD + px(316), y, ContentW() - px(316), px(32), 9, False, CLR_MUTED_FG

    tiles = Array("tileMsgs", "tileRate", "tileOwed", "tileOverdue")
    labels = Array("Messages received", "Acknowledged", "Acknowledgements owed", "Overdue messages")
    glyphs = Array("E715", "E73E", "E823", "E7BA")
    tw = (ContentW() - 3 * px(12)) \ 4
    For i = 0 To 3
        x = PAGE_PAD + i * (tw + px(12))
        StatTile CStr(tiles(i)), CStr(labels(i)), "", x, px(128), IIf(i = 3, ContentW() - 3 * (tw + px(12)), tw), px(80), _
            glyph:=CStr(glyphs(i)), fmt:=IIf(i = 1, "0%", "#,##0")
    Next

    ' Accountability, and the picked message's acknowledgements under it.
    y = CARD_Y: h = PAGE_H - y - px(16): mw = MainW()
    Card "Ack", "Acknowledgement accountability", PAGE_PAD, y, mw, h, , 2, 2
    Set c = Btn("btnOpen", "Open message", PAGE_PAD + mw - px(16) - px(120), y + px(8), px(120), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Read the picked message (or double-click its row)"
    GridHeaders "hdrAck", RptAckCols(), PAGE_PAD + px(16), y + px(40), w:=mw - px(32), sortSf:="sfAck"
    Set c = SubF("sfAck", "subGridRptAck", PAGE_PAD + px(16), y + px(62), mw - px(32), px(150)): Anchor c, 2, 0
    EmptyState "emptyAck", PAGE_PAD + px(16), y + px(70), mw - px(32), "E73E", "Nothing to acknowledge in this period", _
        "Messages that must be acknowledged show here. Pick a longer period."
    Set c = Lbl("lblDet", "", PAGE_PAD + px(16), y + px(220), mw - px(32), px(16), 8, True, CLR_MUTED_FG): Anchor c, 2, 0
    GridHeaders "hdrDet", RptDetCols(), PAGE_PAD + px(16), y + px(240), w:=mw - px(32), sortSf:="sfDet"
    Set c = SubF("sfDet", "subGridRptAckDet", PAGE_PAD + px(16), y + px(262), mw - px(32), h - px(278)): Anchor c, 2, 2

    ' Traffic measures.
    x = PAGE_PAD + mw + px(12)
    Card "Traffic", "Traffic measures", x, y, SIDE_W, h, , 1, 2
    k = BuiltNames()
    Segmented "segMeasure", "", x + px(16), y + px(40), SIDE_W - px(32), _
        "prec|Precedence;type|Type;class|Class;life|Lifecycle", px(32), "MeasureChanged"
    GridHeaders "hdrTraffic", RptTrafficCols(), x + px(16), y + px(84), w:=SIDE_W - px(32)
    AnchorBuilt k, 1, 0                 ' the switch and the captions move with the card
    Set c = SubF("sfTraffic", "subGridRptTraffic", x + px(16), y + px(106), SIDE_W - px(32), h - px(122)): Anchor c, 1, 2
    k = BuiltNames()
    EmptyState "emptyTraffic", x + px(16), y + px(114), SIDE_W - px(32), "E9D2", "No traffic in this period", _
        "Pick a longer period."
    AnchorBuilt k, 1, 0
    SaveForm "pgReports"
End Sub
