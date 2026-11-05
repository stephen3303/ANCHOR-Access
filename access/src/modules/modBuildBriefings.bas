Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Briefings page (/briefings and /briefings/generated): the builder and the generated briefing.
'
' The saved briefings on the left; the builder on the right: title and
' period, the filters with "Add matching messages", the pick list the
' creator ticks and notes, the overall notes, Save and Print briefing.
' The generated briefing is the report rptBriefing (Build_rptBriefing,
' which modBuild.BuildAppDialogs calls with the dialogs, after the pages).
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, FieldLbl), modBriefings (BRIEFING_REPORT); the kit's
' modUIKit.

Private Const LIST_W As Long = 4200             ' the saved briefings card, 280 px

Private Function BuildX() As Long
    BuildX = PAGE_PAD + LIST_W + px(12)
End Function

Private Function BuildW() As Long
    BuildW = ContentW() - LIST_W - px(12)
End Function

' The Briefings page and the tables it hosts (modBuild.BuildPages).
Public Sub BuildBriefingsPage()
    BuildBriefingsGrids
    Build_pgBriefings
End Sub

' The page's tables, before the page.
Public Sub BuildBriefingsGrids()
    BuildGrid "subGridBrfSaved", BrfSavedCols(), LIST_W - px(32)
    BuildGrid "subGridBrfPick", BrfPickCols(), BuildW() - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function BrfSavedCols() As Variant
    BrfSavedCols = Array("created_txt|Created|80|m|CreatedOn", "brief_title|Briefing|168|b")
End Function

Private Function BrfPickCols() As Variant
    BrfPickCols = Array("tick|Brief|48|cb", "ref|Reference|132|b", "dtg|DTG|112|m", "prec|Precedence|96|p|prec_rank", _
        "subject_txt|Subject|160|", "item_note|Note|120|m")
End Function

' The saved briefings list and New briefing; the builder card: title, From
' and To dates, the precedence and type filters (their rows set in code)
' and the community text with "Add matching messages"; the pick list with
' Tick / untick and Note; the briefing's notes; Save and Print briefing.
Public Sub Build_pgBriefings()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, bw As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Briefings", "Build a briefing from the traffic of a period: add the messages that match, tick the ones " & _
        "to brief, add a note to any of them, save it and print it. A message without a note is briefed by the first " & _
        "paragraph of its official text, never by its AI summary.", PAGE_PAD
    Set c = Btn("btnNew", "New briefing", PAGE_W - PAGE_PAD - px(128), px(24), px(128), px(32), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Start a new briefing (Ctrl+N)"

    ' Saved briefings.
    y = px(84): h = PAGE_H - y - px(16)
    Card "Saved", "Saved briefings", PAGE_PAD, y, LIST_W, h, , 0, 2
    GridHeaders "hdrSaved", BrfSavedCols(), PAGE_PAD + px(16), y + px(40), w:=LIST_W - px(32), sortSf:="sfSaved"
    Set c = SubF("sfSaved", "subGridBrfSaved", PAGE_PAD + px(16), y + px(62), LIST_W - px(32), h - px(78)): Anchor c, 0, 2
    EmptyState "emptySaved", PAGE_PAD + px(16), y + px(70), LIST_W - px(32), "E8A5", "No briefings yet", _
        "Briefings you save show here."

    ' The builder.
    x = BuildX(): bw = BuildW(): fx = x + px(16): fw = bw - px(32)
    Card "Build", "New briefing", x, y, bw, h, , 2, 2
    FieldLbl "lblTitleF", "TITLE", fx, y + px(40), px(300)
    Set c = Txt("txtTitle", "", fx, y + px(58), px(300), px(30)): Ev c, "AfterUpdate"
    FieldLbl "lblFrom", "FROM", fx + px(316), y + px(40), DATEFIELD_W
    Set c = DateBox("dtFrom", "", fx + px(316), y + px(58), px(30)): Ev c, "AfterUpdate"
    Ev F.Controls("cal_dtFrom"), "OnClick"     ' the page handles it: a picked date is a change to save
    FieldLbl "lblTo", "TO", fx + px(482), y + px(40), DATEFIELD_W
    Set c = DateBox("dtTo", "", fx + px(482), y + px(58), px(30)): Ev c, "AfterUpdate"
    Ev F.Controls("cal_dtTo"), "OnClick"

    FieldLbl "lblPrec", "PRECEDENCE", fx, y + px(100), px(132)
    Set c = Cbo("cboPrec", "", "", fx, y + px(118), px(132), px(30), 2, "0;" & px(132))
    c.RowSourceType = "Value List"          ' the rows are set in code (BrfPrecedenceList)
    c.ControlTipText = "Only messages of this precedence"
    FieldLbl "lblType", "MESSAGE TYPE", fx + px(140), y + px(100), px(172)
    Set c = Cbo("cboType", "", "", fx + px(140), y + px(118), px(172), px(30), 2, "0;" & px(172))
    c.RowSourceType = "Value List"          ' the rows are set in code, from tblMessageTypes (BrfTypeList)
    c.ControlTipText = "Only messages of this type (the types an admin keeps in the message type list)"
    FieldLbl "lblComm", "COMMUNITY", fx + px(320), y + px(100), px(180)
    Set c = Txt("txtComm", "", fx + px(320), y + px(118), px(180), px(30))
    c.ControlTipText = "Only messages for a community whose name has this text (Safety, Readiness ...); blank for any"
    Set c = Btn("btnAdd", "Add matching messages", fx + px(508), y + px(118), fw - px(508), px(30), "outline")
    c.ControlTipText = "Add the period's messages that match the filters to the list below, not ticked"

    FieldLbl "lblPick", "MESSAGES TO BRIEF", fx, y + px(160), px(160)
    Set c = Lbl("lblCount", "", fx + px(160), y + px(158), fw - px(160) - px(208), px(20), 9, False, CLR_MUTED_FG): Anchor c, 2, 0
    Set c = Btn("btnTick", "Tick / untick", fx + fw - px(200), y + px(154), px(112), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Brief the picked message, or stop briefing it (or double-click its row)"
    Set c = Btn("btnNote", "Note...", fx + fw - px(80), y + px(154), px(80), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Write what the briefing says about the picked message"
    GridHeaders "hdrPick", BrfPickCols(), fx, y + px(190), w:=fw, sortSf:="sfPick"
    Set c = SubF("sfPick", "subGridBrfPick", fx, y + px(212), fw, h - px(350)): Anchor c, 2, 2
    EmptyState "emptyPick", fx, y + px(220), fw, "E8A5", "No messages listed", _
        "Set the period and the filters, then Add matching messages."

    Set c = Lbl("lblNotes", "NOTES ON THE WHOLE BRIEFING", fx, y + h - px(130), fw, px(16), 8, True, CLR_MUTED_FG): Anchor c, 0, 1
    Set c = Txt("txtNotes", "", fx, y + h - px(112), fw, px(56))
    c.EnterKeyBehavior = True
    c.ScrollBars = 2
    Ev c, "AfterUpdate"
    Anchor c, 2, 1
    Set c = Btn("btnSave", "Save", fx, y + h - px(48), px(88), px(32)): Anchor c, 0, 1
    c.ControlTipText = "Save the briefing with its ticked messages (Ctrl+S)"
    Set c = Btn("btnPrint", "Print briefing", fx + px(96), y + h - px(48), px(128), px(32), "outline"): Anchor c, 0, 1
    c.ControlTipText = "Open the saved briefing to print"
    SaveForm "pgBriefings"
End Sub

' ---------------------------------------------------------------------------
' rptBriefing (the generated briefing)
' ---------------------------------------------------------------------------

' The printed briefing, opened by modBriefings.BrfPrint with
' "BriefingID=<id>": the classification line at the top and bottom of each
' page (the highest classification of its messages, BrfMarking), then the
' title, the period, who prepared it and its notes; then each message in
' briefing order: reference, DTG, precedence and classification, the
' subject, and the creator's note or the first paragraph of the official
' text (BrfItemText), never the AI summary. Page numbers at the bottom.
' Build it after the pages (modBuild.BuildAppDialogs).
Public Sub Build_rptBriefing()
    Dim c As Control, w As Long, ghSec As Integer
    w = REPORT_W
    NewReport "SELECT b.BriefingID, b.Title, b.PersonName, b.CreatedOn, b.PeriodFrom, b.PeriodTo, b.Notes, i.SortOrder, " & _
        "i.ItemNote, m.MessageID, m.Reference, m.Dtg, m.Subject, m.Precedence, m.Classification FROM (tblBriefings AS b " & _
        "INNER JOIN tblBriefingItems AS i ON b.BriefingID=i.BriefingID) INNER JOIN tblMessages AS m ON i.MessageID=m.MessageID", w
    F.HasModule = False
    F.Caption = "Briefing"
    ' Grouped on the briefing, for its heading, and sorted in briefing order.
    CreateGroupLevel F.Name, "BriefingID", True, False
    CreateGroupLevel F.Name, "SortOrder", False, False
    ghSec = acGroupLevel1Header

    F.Section(acPageHeader).Height = px(32)
    Set c = RO("txtMarkTop", "=BrfMarking([BriefingID])", 0, px(4), w, px(20), 10, True, PRINT_FG, acPageHeader)
    c.TextAlign = 2

    F.Section(ghSec).Height = px(150)
    F.Section(ghSec).CanGrow = True
    Set c = RO("txtTitle", "Title", 0, px(8), w, px(32), 16, True, PRINT_TITLE, ghSec)
    c.CanGrow = True
    RO "txtPeriod", "=""Period: "" & Format([PeriodFrom],""dd mmm yyyy"") & "" to "" & Format([PeriodTo],""dd mmm yyyy"")", _
        0, px(44), w, px(20), 10, False, PRINT_FG, ghSec
    RO "txtBy", "=""Prepared by "" & Nz([PersonName],""(not recorded)"") & "" on "" & Format([CreatedOn],""dd mmm yyyy hh:nn"")", _
        0, px(66), w, px(20), 10, False, PRINT_MUTED, ghSec
    Set c = RO("txtNotes", "=IIf(IsNull([Notes]),"""",""Notes: "" & [Notes])", 0, px(94), w, px(20), 10, False, PRINT_FG, ghSec)
    c.CanGrow = True
    c.CanShrink = True
    Box "boxHeadLine", 0, px(140), w, px(1), PRINT_FG, -1, ghSec

    F.Section(acDetail).Height = px(84)
    F.Section(acDetail).CanGrow = True
    F.Section(acDetail).BackColor = PRINT_PAPER
    F.Section(acDetail).AlternateBackColor = PRINT_PAPER
    RO "txtItemHead", "=[Reference] & ""   "" & Nz([Dtg],"""") & ""   "" & [Precedence] & ""   "" & ClassificationMarking([Classification])", _
        0, px(10), w, px(18), 9, True, PRINT_MUTED, acDetail
    Set c = RO("txtSubject", "Subject", 0, px(30), w, px(20), 11, True, PRINT_FG, acDetail)
    c.CanGrow = True
    Set c = RO("txtItem", "=BrfItemText([MessageID],[ItemNote])", 0, px(54), w, px(20), 10, False, PRINT_FG, acDetail)
    c.CanGrow = True

    F.Section(acPageFooter).Height = px(48)
    Set c = RO("txtMarkBot", "=BrfMarking([BriefingID])", 0, px(4), w, px(20), 10, True, PRINT_FG, acPageFooter)
    c.TextAlign = 2
    Set c = RO("txtPage", "=""Page "" & [Page] & "" of "" & [Pages]", 0, px(26), w, px(18), 8, False, PRINT_MUTED, acPageFooter)
    c.TextAlign = 3
    SaveForm BRIEFING_REPORT
End Sub
