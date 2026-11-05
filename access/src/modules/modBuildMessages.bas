Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Messages page (/messages: the list with its
' filters and preview), the Message page (/messages/[id]: the official text,
' the AI summary apart, integrity, lifecycle and Acknowledge) and the
' printed message, rptMessage (/messages/[id]/print). modBuild.BuildPages
' calls BuildMessagesPages, which builds the tables first, then the pages
' and the report.
'
' Owner: the message pages' build task. Depends on: modBuild (PAGE_W,
' PAGE_H, PAGE_PAD, ContentW, FieldLbl, NamesNow, GroupSince); modUI
' (BadgeColors); the kit's modUIKit.

Private Const PREVIEW_W As Long = 4800          ' the Messages page's preview card, 320 px
Private Const MSG_SIDE_W As Long = 5640         ' the Message page's tabbed card, 376 px

' Pill variants, in the order a pill's conditions are added (modUI.BadgeColors).
Private Const PILL_VARIANTS As String = "success,warning,info,destructive,default,outline,secondary"

Private Function MsgListW() As Long
    MsgListW = ContentW() - PREVIEW_W - px(12)
End Function

Private Function MsgMainW() As Long
    MsgMainW = ContentW() - MSG_SIDE_W - px(12)
End Function

' The pages, their tables first, and the printed message.
Public Sub BuildMessagesPages()
    BuildMessagesGrids
    Build_pgMessages
    Build_pgMessage
    Build_rptMessage
End Sub

' The dialogs the message pages open (modBuild.BuildAppDialogs): none; the
' printed message is a report, built with the pages.
Public Sub BuildMessagesDialogs()
End Sub

' Every table the message pages show, before the forms that host them.
Public Sub BuildMessagesGrids()
    Dim sw As Long
    sw = MSG_SIDE_W - px(32)
    BuildGrid "subGridMessages", MsgListCols(), MsgListW() - px(32)
    BuildGrid "subGridMsgAddr", MsgAddrCols(), sw
    BuildGrid "subGridMsgAtt", MsgAttCols(), sw
    BuildGrid "subGridMsgAcks", MsgAckCols(), sw
    BuildGrid "subGridMsgRelated", MsgRelatedCols(), sw, 2
End Sub

' "field|caption|width px|options[|sort field]"
Private Function MsgListCols() As Variant
    MsgListCols = Array("dot||20|c|read_state", "subject|Subject|200|s", "ref|Reference|128|s", _
        "prec|Precedence|92|p|prec_rank", "lifecycle|Status|100|p", "ack_due|Ack|72|p")
End Function

Private Function MsgAddrCols() As Variant
    MsgAddrCols = Array("addr_name|Addressee|208|", "addr_role|Role|56|m", "addr_kind|Kind|64|m")
End Function

Private Function MsgAttCols() As Variant
    MsgAttCols = Array("att_name|File|248|", "att_size|Size|80|mr")
End Function

Private Function MsgAckCols() As Variant
    MsgAckCols = Array("cmd_name|Command|152|", "ack_state|Status|100|p", "ack_at|Received|80|m|first_at")
End Function

Private Function MsgRelatedCols() As Variant
    MsgRelatedCols = Array("lifecycle|Status|104|p", "item_txt|Message|224|")
End Function

' A status pill on a single form: an unbound text box the page sets, drawn
' as a rounded pill in StatusVariant's colors (as BuildGrid's "p" columns).
' hAnchor anchors it (see Anchor) before its pill image is made, which
' copies the anchoring. Returns the text box.
Private Function PillBox(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, _
        Optional ByVal hAnchor As Integer = 0) As Control
    Dim c As Control, v As Variant, col As Variant, first As Integer
    Set c = RO(nm, "", x, y, w, px(20), 8, True, CLR_FG)
    c.TextAlign = 2
    Anchor c, hAnchor, 0
    first = c.FormatConditions.Count
    For Each v In Split(PILL_VARIANTS, ",")
        col = BadgeColors(CStr(v))
        CF c, "StatusVariant([" & nm & "])=""" & v & """", col(0), col(1)
    Next
    RoundCF c, first
    Set PillBox = c
End Function

' A caption and its value on one line, for the Message page's details: the
' muted caption lbl<nm> 112 px wide and the value txt<nm> beside it.
Private Function KeyValue(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, _
        Optional ByVal h As Long = 330) As Control
    Lbl "lbl" & nm, cap, x, y, px(112), px(20), 9, False, CLR_MUTED_FG
    Set KeyValue = RO("txt" & nm, "", x + px(112), y, w - px(112), h, 9)
End Function

' ---------------------------------------------------------------------------
' Messages
' ---------------------------------------------------------------------------

' The search box, the All / Unread / Ack due switch and the precedence,
' type and classification combos over every message, newest first; the
' picked one's preview on the right with Open. Mark read, Mark unread and
' Mark all read under the list.
Public Sub Build_pgMessages()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, lw As Long, btnY As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnTimer"             ' OnTimer: the list loads once the page is open
    PageHeader "Messages", "Every message ANCHOR received, newest first. Pick one to see its preview; double-click it, " & _
        "or click Open, to read it in full. The dot marks the messages you haven't read.", PAGE_PAD
    y = px(84)
    SearchBox "txtSearch", PAGE_PAD, y, px(240), "SearchFor", "Search messages...", px(32)
    x = PAGE_PAD + px(252)
    Segmented "segShow", "", x, y, px(240), "all|All;unread|Unread;ack|Ack due", px(32), "ShowChanged"
    x = x + px(252)
    Set c = Cbo("cboPrec", "", MSG_ALL_PREC & ";FLASH;IMMEDIATE;PRIORITY;ROUTINE", x, y, px(128), px(32))
    c.ControlTipText = "Only messages of this precedence": Ev c, "AfterUpdate", "OnGotFocus"
    x = x + px(136)
    Set c = Cbo("cboType", "", MSG_ALL_TYPES, x, y, px(160), px(32))
    c.ControlTipText = "Only messages of this type (the types an admin keeps)": Ev c, "AfterUpdate", "OnGotFocus"
    x = x + px(168)
    Set c = Cbo("cboClass", "", MSG_ALL_CLASSES & ";UNCLASSIFIED;CUI;CONFIDENTIAL;SECRET", x, y, px(140), px(32))
    c.ControlTipText = "Only messages of this classification": Ev c, "AfterUpdate", "OnGotFocus"

    y = px(128): h = PAGE_H - y - px(16): lw = MsgListW(): btnY = y + h - px(46)
    Card "List", "Messages", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrList", MsgListCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfList"
    Set c = SubF("sfList", "subGridMessages", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(124)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), lw - px(32), "E715", "No messages match", _
        "No message fits the search and the filters.", "Clear filters"
    Set c = Btn("btnMarkRead", "Mark read", PAGE_PAD + px(16), btnY, px(96), px(30), "outline"): Anchor c, 0, 1
    c.ControlTipText = "Mark the picked message read"
    Set c = Btn("btnMarkUnread", "Mark unread", PAGE_PAD + px(120), btnY, px(104), px(30), "outline"): Anchor c, 0, 1
    c.ControlTipText = "Mark the picked message unread again"
    Set c = Btn("btnMarkAll", "Mark all read", PAGE_PAD + px(232), btnY, px(112), px(30), "outline"): Anchor c, 0, 1
    c.ControlTipText = "Mark every message read"

    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = PREVIEW_W - px(32)
    Card "Preview", "Preview", x, y, PREVIEW_W, h, , 1, 2
    Set c = Lbl("lblPvSubject", "", fx, y + px(40), fw, px(44), 11, True): Anchor c, 1, 0
    Set c = Lbl("lblPvRef", "", fx, y + px(88), fw, px(18), 9, False, CLR_MUTED_FG): Anchor c, 1, 0
    Set c = Lbl("lblPvFrom", "", fx, y + px(106), fw, px(18), 9, False, CLR_MUTED_FG): Anchor c, 1, 0
    PillBox "pvPrec", fx, y + px(132), px(84), 1
    PillBox "pvClass", fx + px(88), y + px(132), px(104), 1
    PillBox "pvLife", fx + px(196), y + px(132), fw - px(196), 1
    Set c = Lbl("lblPvNote", "", fx, y + px(160), fw, px(32), 8, False, CLR_MUTED_FG): Anchor c, 1, 0
    c.TopMargin = 0
    FieldLbl "capPvBody", "OFFICIAL TEXT (FIRST LINES)", fx, y + px(196), fw
    Anchor F.Controls("capPvBody"), 1, 0
    Set c = RO("txtPvBody", "", fx, y + px(214), fw, btnY - px(8) - (y + px(214)), 9)
    c.BorderStyle = 1: c.BorderColor = CLR_BORDER: c.ScrollBars = 2: Anchor c, 1, 2
    Set c = Btn("btnOpen", "Open", fx, btnY, px(88), px(30)): Anchor c, 1, 1
    c.ControlTipText = "Read the picked message in full (or double-click its row)"
    Set c = Lbl("lblDetailEmpty", "Pick a message in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2: Anchor c, 1, 0
    SaveForm "pgMessages"
End Sub

' ---------------------------------------------------------------------------
' Message
' ---------------------------------------------------------------------------

' One message: the subject and its pills over the lifecycle banner and the
' acknowledgement bar (with Acknowledge); the official text, with a switch
' to the raw text as received, and the AI summary in its own card under it;
' on the right four tabs: Details (type, class, integrity, the action item),
' Addressees (with the attachments' names), Acks (what the addressees of
' this command's own traffic acknowledged) and Related. Back, Mark unread
' and Print at the top.
Public Sub Build_pgMessage()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, mw As Long, sx As Long, before As Object
    Dim ty As Long, k As String
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnTimer"             ' OnTimer: the message loads once the page is open, and a related one once its click has finished
    PageHeader "Message", "The official text exactly as received, its integrity data, its status and your acknowledgement. " & _
        "The AI summary is a reading aid only, shown apart and never the official record.", PAGE_PAD
    Set c = Btn("btnPrint", "Print", PAGE_W - PAGE_PAD - px(80), px(24), px(80), px(32), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Preview the official rendition with its classification banners and integrity block, to print"
    Set c = Btn("btnUnread", "Mark unread", PAGE_W - PAGE_PAD - px(80) - px(8) - px(112), px(24), px(112), px(32), "outline")
    Anchor c, 1, 0: c.ControlTipText = "Mark this message unread again"
    Set c = Btn("btnBack", "Back to Messages", PAGE_W - PAGE_PAD - px(80) - px(8) - px(112) - px(8) - px(144), px(24), px(144), _
        px(32), "outline")
    Anchor c, 1, 0: c.ControlTipText = "Go back to the list of messages"

    mw = MsgMainW(): sx = PAGE_PAD + mw + px(12)
    Set c = Lbl("lblSubject", "", PAGE_PAD, px(84), ContentW(), px(26), 13, True): Anchor c, 2, 0
    PillBox "pilClass", PAGE_PAD, px(116), px(112)
    PillBox "pilPrec", PAGE_PAD + px(120), px(116), px(96)
    PillBox "pilLife", PAGE_PAD + px(224), px(116), px(104)
    Set c = Lbl("lblMarking", "", PAGE_PAD + px(340), px(116), ContentW() - px(340), px(20), 8, True, CLR_DANGER_TEXT)
    Anchor c, 2, 0

    ' The lifecycle banner, with the replacing or related message's link at
    ' its right, and the acknowledgement bar, with Acknowledge at its right.
    y = px(148)
    Alert "alLife", PAGE_PAD, y, mw, px(52)
    Anchor F.Controls("alLife"), 2, 0
    Anchor F.Controls("alLife_title"), 2, 0
    Set c = F.Controls("alLife_msg"): c.Width = mw - px(44) - px(16) - px(168): Anchor c, 2, 0
    Set c = Link("lnkLife", "", PAGE_PAD + mw - px(176), y + px(16), px(160), px(20), "", 9, CLR_LINK): Anchor c, 1, 0
    c.TextAlign = 3: c.FontUnderline = True: c.OnClick = "[Event Procedure]"
    c.ControlTipText = "Open that message"
    k = BuiltNames()
    Alert "alAck", sx, y, MSG_SIDE_W, px(52)
    F.Controls("alAck_msg").Width = MSG_SIDE_W - px(44) - px(16) - px(124)
    Set c = Btn("btnAck", "Acknowledge", sx + MSG_SIDE_W - px(12) - px(112), y + px(12), px(112), px(28), "destructive")
    c.ControlTipText = "Record that you received and read this message"
    AnchorBuilt k, 1, 0

    ' The official text and the raw text in one place, and the AI summary apart.
    y = px(212): fx = PAGE_PAD + px(16): fw = mw - px(32)
    Card "Text", "Official text", PAGE_PAD, y, mw, px(292), , 2, 2
    k = BuiltNames()
    Segmented "segText", "", PAGE_PAD + mw - px(16) - px(208), y + px(8), px(208), "fmt|As formatted;raw|Raw as received", _
        px(28), "TextViewChanged"
    AnchorBuilt k, 1, 0
    Set c = RO("txtBody", "", fx, y + px(44), fw, px(292) - px(60), 10)
    c.BorderStyle = 1: c.BorderColor = CLR_BORDER: c.ScrollBars = 2: Anchor c, 2, 2
    Set c = RO("txtRaw", "", fx, y + px(44), fw, px(292) - px(60), 9)
    c.FontName = "Consolas": c.BorderStyle = 1: c.BorderColor = CLR_BORDER: c.ScrollBars = 2: Anchor c, 2, 2
    c.Visible = False
    y = px(516)
    Card "Ai", "AI summary - Not authoritative", PAGE_PAD, y, mw, PAGE_H - px(16) - y, , 2, 1
    Set c = RO("txtAi", "", fx, y + px(36), fw, px(38), 9): Anchor c, 2, 1
    c.ScrollBars = 2
    Set c = Lbl("lblAiNote", "Generated from the official text. It may be wrong: always act on the official text.", _
        fx, y + px(76), fw, px(16), 8, False, CLR_MUTED_FG)
    Anchor c, 2, 1

    ' The tabbed card on the right.
    y = px(212): h = PAGE_H - px(16) - y: fx = sx + px(16): fw = MSG_SIDE_W - px(32)
    Card "Side", "", sx, y, MSG_SIDE_W, h, , 1, 2
    ' Everything in the card moves with its right edge (AnchorBuilt below);
    ' its tables also stretch down with it.
    k = BuiltNames()
    TabStrip fx, y + px(10), fw, "TabPick", "d|Details", "a|Addressees", "k|Acks", "r|Related"
    ty = y + px(52)

    Set before = NamesNow()
    FieldLbl "capMsg", "MESSAGE", fx, ty, fw
    KeyValue "Type", "Type", fx, ty + px(18), fw
    KeyValue "Class", "Class", fx, ty + px(40), fw
    KeyValue "Received", "Received", fx, ty + px(62), fw
    FieldLbl "capIntegrity", "INTEGRITY (RECORDED AT INGESTION)", fx, ty + px(92), fw
    KeyValue "C2oix", "C2OIX id", fx, ty + px(110), fw
    Set c = KeyValue("Sha", "SHA-256", fx, ty + px(132), fw, px(30))
    c.FontName = "Consolas": c.FontSize = 8
    KeyValue "Size", "Size", fx, ty + px(164), fw
    KeyValue "Retention", "Retention", fx, ty + px(186), fw
    FieldLbl "capAction", "ACTION ITEM", fx, ty + px(216), fw
    Set c = RO("txtAction", "", fx, ty + px(234), fw, y + h - px(16) - (ty + px(234)), 9)
    GroupSince before, "grpD"

    Set before = NamesNow()
    FieldLbl "capAddr", "ADDRESSEES", fx, ty, fw
    GridHeaders "hdrAddr", MsgAddrCols(), fx, ty + px(18), w:=fw
    Set c = SubF("sfAddr", "subGridMsgAddr", fx, ty + px(40), fw, px(136)): c.TabStop = False
    FieldLbl "capDist", "DISTRIBUTION", fx, ty + px(184), fw
    RO "txtDist", "", fx, ty + px(202), fw, px(36), 9
    FieldLbl "capAtt", "ATTACHMENTS (NAMES ONLY)", fx, ty + px(244), fw
    GridHeaders "hdrAtt", MsgAttCols(), fx, ty + px(262), w:=fw
    Set c = SubF("sfAtt", "subGridMsgAtt", fx, ty + px(284), fw, y + h - px(16) - (ty + px(284))): c.TabStop = False
    GroupSince before, "grpA"
    Set c = Lbl("lblAttNone", "No attachments.", fx, ty + px(290), fw, px(20), 9, False, CLR_MUTED_FG)
    c.Visible = False

    Set before = NamesNow()
    Set c = Lbl("lblAckSum", "", fx, ty, fw, px(52), 9, False, CLR_FG)     ' three lines
    c.TopMargin = 0
    GridHeaders "hdrAcks", MsgAckCols(), fx, ty + px(60), w:=fw, sortSf:="sfAcks"
    Set c = SubF("sfAcks", "subGridMsgAcks", fx, ty + px(82), fw, y + h - px(16) - (ty + px(82))): c.TabStop = False
    GroupSince before, "grpK"

    Set before = NamesNow()
    Set c = Lbl("lblRelHint", "Click a message to open it.", fx, ty, fw, px(18), 9, False, CLR_MUTED_FG)
    GridHeaders "hdrRel", MsgRelatedCols(), fx, ty + px(24), w:=fw
    Set c = SubF("sfRel", "subGridMsgRelated", fx, ty + px(46), fw, y + h - px(16) - (ty + px(46))): c.TabStop = False
    GroupSince before, "grpR"
    Set c = Lbl("lblRelNone", "ANCHOR holds no message this one names.", fx, ty + px(60), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2: c.Visible = False
    AnchorBuilt k, 1, 0
    Anchor F.Controls("txtAction"), 1, 2
    Anchor F.Controls("sfAtt"), 1, 2
    Anchor F.Controls("sfAcks"), 1, 2
    Anchor F.Controls("sfRel"), 1, 2

    ' No message to show (none picked, or ANCHOR doesn't hold it).
    EmptyState "emptyMsg", PAGE_PAD, px(200), ContentW(), "E715", "No message to show", _
        "ANCHOR does not hold this message. It may not have been ingested yet.", "Back to Messages"
    SaveForm "pgMessage"
End Sub

' ---------------------------------------------------------------------------
' The printed message
' ---------------------------------------------------------------------------

' The official rendition (the web app's print view): the classification
' marking as a banner at the top and bottom of every page; the heading block
' (precedence and DTG, FM, TO, INFO, MSGID, SUBJ), the official text and the
' integrity block, which grow to their text; page numbers. Its code-behind
' sets the record source to the message OpenArgs names ("id=12").
Public Sub Build_rptMessage()
    Dim c As Control
    NewReport "", REPORT_W
    F.OnOpen = "[Event Procedure]"
    F.OnNoData = "[Event Procedure]"
    F.Caption = "Message"
    F.Section(acPageHeader).Height = px(40)
    F.Section(acPageFooter).Height = px(56)
    F.Section(acDetail).Height = px(120)
    F.Section(acDetail).CanGrow = True
    F.Section(acPageHeader).BackColor = PRINT_PAPER
    F.Section(acPageFooter).BackColor = PRINT_PAPER
    F.Section(acDetail).BackColor = PRINT_PAPER
    F.Section(acDetail).AlternateBackColor = PRINT_PAPER

    Set c = RO("txtBannerTop", "=ClassificationMarking([Classification])", 0, px(4), REPORT_W, px(22), 11, True, PRINT_FG, acPageHeader)
    c.TextAlign = 2
    Rule "ruleTop", 0, px(32), REPORT_W, 0, PRINT_FG, acPageHeader

    Set c = RO("txtHead", "=MessagePrintHead([MessageID])", 0, px(8), REPORT_W, px(20), 10, False, PRINT_FG, acDetail)
    c.FontName = "Consolas": c.CanGrow = True
    Set c = RO("txtBody", "BodyText", 0, px(40), REPORT_W, px(20), 10, False, PRINT_FG, acDetail)
    c.FontName = "Consolas": c.CanGrow = True
    Set c = RO("txtIntegrity", "=MessagePrintIntegrity([MessageID])", 0, px(80), REPORT_W, px(20), 8, False, PRINT_MUTED, acDetail)
    c.CanGrow = True: c.BorderStyle = 1: c.BorderColor = PRINT_MUTED

    Rule "ruleBottom", 0, px(2), REPORT_W, 0, PRINT_FG, acPageFooter
    Set c = RO("txtBannerBottom", "=ClassificationMarking([Classification])", 0, px(8), REPORT_W, px(22), 11, True, PRINT_FG, acPageFooter)
    c.TextAlign = 2
    Set c = RO("txtPage", "=[Reference] & ""   Page "" & [Page] & "" of "" & [Pages]", 0, px(32), REPORT_W, px(18), 8, False, _
        PRINT_MUTED, acPageFooter)
    c.TextAlign = 3
    SaveForm "rptMessage"
End Sub
