Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the drafting pages, which came over from Message
' Generator when ANCHOR took over drafting and release (Stephen,
' 2026-11-05), laid out as the web app's screens:
'   pgDrafting   Drafting & Release (/drafting): the drafts, and the picked
'                one with its Draft, Message (the text as it transmits and
'                the naval format check), Route and Log tabs
'   pgTemplates  Templates (/templates): the templates with their versions,
'                and the picked one with its Header, Text, Fields and
'                Preview tabs
'   pgRouting    Admin: Routing (/admin/routing): workflows, release
'                authority and the command's transmission settings
' and the dialogs they open: dlgNewDraft, dlgTemplateField and dlgRouteStep.
' modBuild.BuildPages calls BuildDraftingPages (the tables first, then the
' pages); modBuild.BuildAppDialogs calls BuildDraftingDialogs.
'
' Owner: drafting and release. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, FieldLbl, NamesNow, GroupSince), modDrafting,
' modTemplates; the kit's modUIKit.

Private Const DRAFT_LIST_W As Long = 5400       ' the Drafting page's list card, 360 px
Private Const TPL_LIST_W As Long = 5400         ' the Templates page's list card, 360 px
Private Const WF_LIST_W As Long = 6000          ' Routing: the workflow list card, 400 px
Private Const RA_LIST_W As Long = 8400          ' Routing: the release authority list card, 560 px

Private Function DraftDetailW() As Long
    DraftDetailW = ContentW() - DRAFT_LIST_W - px(12)
End Function

Private Function TplDetailW() As Long
    TplDetailW = ContentW() - TPL_LIST_W - px(12)
End Function

Private Function WfDetailW() As Long
    WfDetailW = ContentW() - WF_LIST_W - px(12)
End Function

Private Function RaFormW() As Long
    RaFormW = ContentW() - RA_LIST_W - px(12)
End Function

' The three pages, each after the tables it hosts.
Public Sub BuildDraftingPages()
    BuildGrid "subGridDrafts", DraftCols(), DRAFT_LIST_W - px(32)
    BuildGrid "subGridRouteSteps", RouteStepCols(), DraftDetailW() - px(32)
    BuildGrid "subGridRouteLog", RouteLogCols(), DraftDetailW() - px(32)
    BuildGrid "subGridTemplates", TemplateCols(), TPL_LIST_W - px(32)
    BuildGrid "subGridTemplateFields", TemplateFieldCols(), TplDetailW() - px(32)
    BuildGrid "subGridWorkflows", WorkflowCols(), WF_LIST_W - px(32)
    BuildGrid "subGridWorkflowSteps", WorkflowStepCols(), WfDetailW() - px(32)
    BuildGrid "subGridAuthorities", AuthorityCols(), RA_LIST_W - px(32)
    Build_pgDrafting
    Build_pgTemplates
    Build_pgRouting
End Sub

' The dialogs the drafting pages open.
Public Sub BuildDraftingDialogs()
    Build_dlgNewDraft
    Build_dlgTemplateField
    Build_dlgRouteStep
End Sub

' "field|caption|width px|options[|sort field]"
Private Function DraftCols() As Variant
    DraftCols = Array("status|Status|96|p", "prec|Precedence|92|p", "subject|Subject|140|bs")
End Function

Private Function RouteStepCols() As Variant
    RouteStepCols = Array("n|#|32|m", "step_txt|Step|170|b", "holder|Held by|130|", "step_state|Status|112|p", "acted_txt|Acted|144|m")
End Function

Private Function RouteLogCols() As Variant
    RouteLogCols = Array("at_txt|When|104|m|EventAt", "event_txt|What|170|b", "ActorName|Who|120|", "EventComment|Comment|194|m")
End Function

Private Function TemplateCols() As Variant
    TemplateCols = Array("tpl|Template|176|bs|TemplateName", "status|Status|88|p", "fields_n|Fields|64|r")
End Function

Private Function TemplateFieldCols() As Variant
    TemplateFieldCols = Array("FieldName|Field|180|b", "req_txt|Rule|88|m", "max_txt|Max length|96|m", "Hint|Hint for drafters|224|m")
End Function

Private Function WorkflowCols() As Variant
    WorkflowCols = Array("WorkflowName|Workflow|176|b", "state|State|72|p", "matches_txt|Matches|120|m")
End Function

Private Function WorkflowStepCols() As Variant
    WorkflowStepCols = Array("n|#|32|m", "StepTitle|Step|200|b", "action_txt|Action|88|m", "holder|Held by|228|")
End Function

Private Function AuthorityCols() As Variant
    AuthorityCols = Array("holder|Person|150|b", "Designation|Releases as|140|", "scope_txt|Message types|120|m", "role_txt|Role|118|p")
End Function

' An Alert built hidden (Alert builds it shown), so a tab's group leaves it
' out and the page shows it with ShowAlert.
Private Sub HiddenAlert(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim s As Variant
    Alert nm, x, y, w, h
    For Each s In Array("", "_bar", "_icon", "_title", "_msg")
        F.Controls(nm & s).Visible = False
    Next
End Sub

' A multi-line text box: Enter starts a new line, a scroll bar when long.
Private Function Memo(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Control
    Dim c As Control
    Set c = Txt(nm, "", x, y, w, h)
    c.EnterKeyBehavior = True
    c.ScrollBars = 2
    Set Memo = c
End Function

' The message as it transmits: read-only, fixed-width, scrolling.
Private Function MessageBox(ByVal nm As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Control
    Dim c As Control
    Set c = RO(nm, "", x, y, w, h, 9)
    c.FontName = "Consolas"
    c.ScrollBars = 2
    Set MessageBox = c
End Function

' The draft header's pickers (the Drafting page's Draft tab and the
' Templates page's Header tab): type, class, precedence, classification.
Private Sub HeaderPickers(ByVal prefix As String, ByVal x As Long, ByVal y As Long, ByVal w As Long)
    Dim cw As Long
    cw = w - px(180) - px(120) - px(120) - px(24)
    FieldLbl "cap" & prefix & "Type", "MESSAGE TYPE", x, y, px(180)
    Cbo "cbo" & prefix & "Type", "", "", x, y + px(18), px(180), px(28), 2, px(180) & ";" & px(260)
    FieldLbl "cap" & prefix & "Class", "CLASS", x + px(188), y, px(120)
    Cbo "cbo" & prefix & "Class", "", "", x + px(188), y + px(18), px(120), px(28), 2, px(32) & ";" & px(220)
    FieldLbl "cap" & prefix & "Prec", "PRECEDENCE", x + px(316), y, px(120)
    Cbo "cbo" & prefix & "Prec", "", "ROUTINE;PRIORITY;IMMEDIATE;FLASH", x + px(316), y + px(18), px(120), px(28)
    FieldLbl "cap" & prefix & "Classif", "CLASSIFICATION", x + px(444), y, cw
    Cbo "cbo" & prefix & "Classif", "", "UNCLASSIFIED;CUI;CONFIDENTIAL;SECRET", x + px(444), y + px(18), cw, px(28)
End Sub

' ---------------------------------------------------------------------------
' Drafting & Release
' ---------------------------------------------------------------------------

' The search box and the chips over the drafts (the latest change first),
' and the picked one on the right in four tabs: Draft (the header and the
' text, Save, Undo, Withdraw), Message (the text as it transmits and the
' naval format check), Route (where it is, the chain, the comment and the
' action the signed-in person may take) and Log (the routing log). New draft
' is top right.
Public Sub Build_pgDrafting()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, btnY As Long, lw As Long, cy As Long
    Dim before As Object, tw As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Drafting & Release", "Draft outgoing command traffic, route it through its review chain and release it. " & _
        "A released message joins Messages.", PAGE_PAD
    Set c = Btn("btnNew", "New draft", PAGE_W - PAGE_PAD - px(120), px(24), px(120), px(32), "outline")
    c.ControlTipText = "Start a blank draft or one from a template (Ctrl+N)": Anchor c, 1, 0
    SearchBox "txtSearch", PAGE_PAD, px(84), px(260), "SearchFor", "Search drafts...", px(28)
    ChipRow "chip", PAGE_PAD + px(272), px(85), "ChipPick", Split(DRAFT_CHIPS, "|")
    y = px(128): h = PAGE_H - y - px(16): lw = DRAFT_LIST_W
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = DraftDetailW() - px(32): btnY = y + h - px(16) - px(30)

    Card "List", "Drafts", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrList", DraftCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfList", _
        filterSf:="sfList", onFilter:="ListFiltered"
    Set c = SubF("sfList", "subGridDrafts", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), lw - px(32), "E70F", "No drafts yet", _
        "New draft starts one, blank or from a template.", "Show all"

    Card "Detail", "Draft", x, y, DraftDetailW(), h, , 1, 2
    Lbl "lblDraftRef", "", fx + px(64), y + px(14), fw - px(64), px(20), 9, False, CLR_MUTED_FG
    TabStrip fx, y + px(40), fw, "TabPick", "draft|Draft", "message|Message", "route|Route", "log|Log"
    cy = y + px(84)

    ' -- Draft --
    Set before = NamesNow()
    FieldLbl "capSubject", "SUBJECT", fx, cy, fw
    Txt "txtSubject", "", fx, cy + px(18), fw, px(28)
    HeaderPickers "", fx, cy + px(54), fw
    Ev F.Controls("cboClassif"), "AfterUpdate"     ' the demo marking follows the pick
    Lbl "lblMarking", "", fx, cy + px(102), fw, px(16), 8, False, CLR_MUTED_FG
    tw = (fw - px(188) - px(16)) \ 2
    FieldLbl "capOriginator", "FROM (FM)", fx, cy + px(122), px(188)
    Txt "txtOriginator", "", fx, cy + px(140), px(188), px(28)
    FieldLbl "capTo", "TO (ACTION; SEPARATE WITH ;)", fx + px(196), cy + px(122), tw
    Txt "txtTo", "", fx + px(196), cy + px(140), tw, px(28)
    FieldLbl "capInfo", "INFO", fx + px(204) + tw, cy + px(122), tw
    Txt "txtInfo", "", fx + px(204) + tw, cy + px(140), tw, px(28)
    FieldLbl "capParas", "TEXT (A BLANK LINE STARTS THE NEXT PARAGRAPH; THEY ARE NUMBERED FOR YOU)", fx, cy + px(176), fw
    Memo "txtParas", fx, cy + px(194), fw, btnY - px(8) - (cy + px(194))
    Set c = Btn("btnSave", "Save", fx, btnY, px(88), px(30))
    c.ControlTipText = "Save the draft (Ctrl+S)"
    Btn "btnUndo", "Undo", fx + px(96), btnY, px(88), px(30), "outline"
    Lbl "lblEditNote", "", fx + px(192), btnY, fw - px(192) - px(112), px(30), 9, False, CLR_MUTED_FG
    Set c = Btn("btnWithdraw", "Withdraw", fx + fw - px(104), btnY, px(104), px(30), "destructive")
    c.ControlTipText = "Withdraw it before release; it is kept for the record"
    GroupSince before, "grpDraft"

    ' -- Message --
    Set before = NamesNow()
    MessageBox "txtPreview", fx, cy, fw, px(248)
    FieldLbl "capIssues", "NAVAL FORMAT CHECK (NTP 3, ACP 121, NAVEDTRA 14226)", fx, cy + px(256), fw
    Set c = RO("txtIssues", "", fx, cy + px(274), fw, btnY - px(8) - (cy + px(274)), 9)
    c.ScrollBars = 2
    Set c = Btn("btnCheck", "Check format", fx, btnY, px(120), px(30), "outline")
    c.ControlTipText = "Run the naval format check on the text as it would transmit if released now"
    Set c = Btn("btnCopy", "Copy text", fx + px(128), btnY, px(96), px(30), "outline")
    c.ControlTipText = "Copy the message as shown"
    Lbl "lblPreviewNote", "", fx + px(232), btnY, fw - px(232), px(30), 8, False, CLR_MUTED_FG
    GroupSince before, "grpMessage"

    ' -- Route --
    Set before = NamesNow()
    Lbl "lblRouteState", "", fx, cy, fw, px(20), 10, True
    Lbl "lblRouteInfo", "", fx, cy + px(20), fw, px(34), 9, False, CLR_MUTED_FG
    FieldLbl "capSteps", "CHAIN", fx, cy + px(118), fw
    GridHeaders "hdrSteps", RouteStepCols(), fx, cy + px(136), w:=fw
    Set c = SubF("sfSteps", "subGridRouteSteps", fx, cy + px(158), fw, px(120))
    c.TabStop = False
    FieldLbl "capComment", "COMMENT (REQUIRED TO RETURN)", fx, cy + px(286), fw
    Memo "txtComment", fx, cy + px(304), fw, btnY - px(8) - (cy + px(304))
    Set c = Btn("btnSubmit", "Submit for routing", fx, btnY, px(150), px(30))
    c.ControlTipText = "Send it to the first step of the chain that matches its type and originator"
    Set c = Btn("btnEndorse", "Endorse", fx, btnY, px(150), px(30))
    c.ControlTipText = "Pass it on to the next step, with your comment"
    Set c = Btn("btnRelease", "Release", fx, btnY, px(150), px(30))
    c.ControlTipText = "Release it: assign the DTG and sequence, freeze the text and add it to Messages"
    Set c = Btn("btnReturn", "Return", fx + px(158), btnY, px(100), px(30), "outline")
    c.ControlTipText = "Send it back to its drafter for correction, with your comment"
    Set c = Btn("btnOpenMessage", "Open in Messages", fx + fw - px(144), btnY, px(144), px(30), "outline")
    c.ControlTipText = "Open the released message"
    GroupSince before, "grpRoute"
    HiddenAlert "alRoute", fx, cy + px(58), fw, px(52)

    ' -- Log --
    Set before = NamesNow()
    Lbl "lblLogNote", "Every step taken on this draft, newest first. The log is only ever added to.", fx, cy, fw, px(20), 9, False, CLR_MUTED_FG
    GridHeaders "hdrLog", RouteLogCols(), fx, cy + px(24), w:=fw, sortSf:="sfLog"
    Set c = SubF("sfLog", "subGridRouteLog", fx, cy + px(46), fw, y + h - px(16) - (cy + px(46)))
    c.TabStop = False
    GroupSince before, "grpLog"

    Set c = Lbl("lblDetailEmpty", "Pick a draft in the list, or start one with New draft.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    SaveForm "pgDrafting"
End Sub

' ---------------------------------------------------------------------------
' Templates
' ---------------------------------------------------------------------------

' The search box and the chips over every template version, and the picked
' one on the right: its name and where it stands, four tabs (Header, Text,
' Fields with each field's rules, Preview), Save and Undo for an Inactive
' version, and Activate or Retire, New version and Start draft. New
' template is top right (ANCHOR admins).
Public Sub Build_pgTemplates()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, btnY As Long, lw As Long, cy As Long
    Dim before As Object, tw As Long, rx As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Templates", "Reusable messages with {{ fields }} the drafter fills in. Drafts start from a template's active version.", PAGE_PAD
    Set c = Btn("btnNewTemplate", "New template", PAGE_W - PAGE_PAD - px(128), px(24), px(128), px(32), "outline")
    c.ControlTipText = "Add a template, inactive until you activate it (Ctrl+N)": Anchor c, 1, 0
    SearchBox "txtSearch", PAGE_PAD, px(84), px(260), "SearchFor", "Search templates...", px(28)
    ChipRow "chip", PAGE_PAD + px(272), px(85), "ChipPick", Split(TPL_CHIPS, "|")
    y = px(128): h = PAGE_H - y - px(16): lw = TPL_LIST_W
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = TplDetailW() - px(32): btnY = y + h - px(16) - px(30)

    Card "List", "Templates", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrList", TemplateCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfList", _
        filterSf:="sfList", onFilter:="ListFiltered"
    Set c = SubF("sfList", "subGridTemplates", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), lw - px(32), "E8A5", "No templates yet", _
        "An ANCHOR admin adds templates with New template.", "Show all"

    Card "Detail", "Template", x, y, TplDetailW(), h, , 1, 2
    Lbl "lblTplName", "", fx, y + px(40), fw, px(22), 11, True
    Lbl "lblTplSub", "", fx, y + px(62), fw, px(18), 9, False, CLR_MUTED_FG
    TabStrip fx, y + px(88), fw, "TabPick", "header|Header", "text|Text", "fields|Fields", "preview|Preview"
    cy = y + px(132)

    ' -- Header --
    Set before = NamesNow()
    FieldLbl "capName", "NAME", fx, cy, px(240)
    Txt "txtName", "", fx, cy + px(18), px(240), px(28)
    FieldLbl "capDesc", "DESCRIPTION", fx + px(248), cy, fw - px(248)
    Txt "txtDesc", "", fx + px(248), cy + px(18), fw - px(248), px(28)
    HeaderPickers "T", fx, cy + px(54), fw
    tw = (fw - px(8)) \ 2
    FieldLbl "capTTo", "TO (ACTION; SEPARATE WITH ;)", fx, cy + px(108), tw
    Txt "txtTTo", "", fx, cy + px(126), tw, px(28)
    FieldLbl "capTInfo", "INFO", fx + tw + px(8), cy + px(108), tw
    Txt "txtTInfo", "", fx + tw + px(8), cy + px(126), tw, px(28)
    FieldLbl "capTSubject", "SUBJECT", fx, cy + px(162), fw
    Txt "txtTSubject", "", fx, cy + px(180), fw, px(28)
    Lbl "lblTagHint", "Write {{ Name }} wherever the drafter fills something in, in the subject, the addressees or the text. " & _
        "Fixed text carries no classification marking.", fx, cy + px(216), fw, px(34), 9, False, CLR_MUTED_FG
    GroupSince before, "grpHeader"

    ' -- Text --
    Set before = NamesNow()
    FieldLbl "capBody", "TEXT (A BLANK LINE STARTS THE NEXT PARAGRAPH; THEY ARE NUMBERED FOR YOU)", fx, cy, fw
    Memo "txtBody", fx, cy + px(18), fw, btnY - px(12) - (cy + px(18))
    GroupSince before, "grpText"

    ' -- Fields --
    Set before = NamesNow()
    Lbl "lblFieldsNote", "One row per {{ field }}, from the tags as saved. Double-click a row, or Field rules, to change its rules.", _
        fx, cy, fw - px(128), px(32), 9, False, CLR_MUTED_FG
    Set c = Btn("btnFieldRules", "Field rules...", fx + fw - px(120), cy, px(120), px(30), "outline")
    c.ControlTipText = "Required, a maximum length, a default and a hint for the picked field"
    GridHeaders "hdrFields", TemplateFieldCols(), fx, cy + px(40), w:=fw
    Set c = SubF("sfFields", "subGridTemplateFields", fx, cy + px(62), fw, btnY - px(12) - (cy + px(62)))
    c.TabStop = False
    GroupSince before, "grpFields"

    ' -- Preview --
    Set before = NamesNow()
    Lbl "lblPreviewNote", "As a draft from it would transmit, each field as its default or [Name]. The DTG and sequence come at release.", _
        fx, cy, fw, px(20), 9, False, CLR_MUTED_FG
    MessageBox "txtTplPreview", fx, cy + px(24), fw, btnY - px(12) - (cy + px(24))
    GroupSince before, "grpPreview"

    Set c = Btn("btnSave", "Save", fx, btnY, px(88), px(30))
    c.ControlTipText = "Save the template (Ctrl+S)"
    Btn "btnUndo", "Undo", fx + px(96), btnY, px(88), px(30), "outline"
    rx = fx + fw
    Set c = Btn("btnStart", "Start draft", rx - px(112), btnY, px(112), px(30))
    c.ControlTipText = "Start a draft from this template: you fill in each field"
    Set c = Btn("btnNewVersion", "New version", rx - px(232), btnY, px(112), px(30), "outline")
    c.ControlTipText = "Copy this version as the next one, inactive, to edit"
    Set c = Btn("btnActivate", "Activate", rx - px(336), btnY, px(96), px(30), "outline")
    c.ControlTipText = "Make this the version drafts start from; the name's active version is retired"
    Set c = Btn("btnRetire", "Retire", rx - px(336), btnY, px(96), px(30), "outline")
    c.ControlTipText = "Stop new drafts starting from it"
    Set c = Lbl("lblDetailEmpty", "Pick a template in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    SaveForm "pgTemplates"
End Sub

' ---------------------------------------------------------------------------
' Admin: Routing
' ---------------------------------------------------------------------------

' Three tabs: Workflows (the list, and the picked one's name, what it
' matches, its chain and its steps' holders), Release authority (who may
' release, and Designate) and Command settings (the text case, FM line,
' MSGID originator and next sequence number). Admins only (PageAllowed).
Public Sub Build_pgRouting()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, btnY As Long, lw As Long
    Dim before As Object, hw As Long, sx As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Routing", "Admin: the chains drafts follow, who may release them, and how the command's messages are set.", PAGE_PAD
    Set c = Btn("btnBack", "Back to Admin", PAGE_W - PAGE_PAD - px(128), px(24), px(128), px(32), "ghost")
    Anchor c, 1, 0
    TabStrip PAGE_PAD, px(84), ContentW(), "TabPick", "wf|Workflows", "ra|Release authority", "cs|Command settings"
    y = px(128): h = PAGE_H - y - px(16)

    ' -- Workflows --
    Set before = NamesNow()
    lw = WF_LIST_W
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = WfDetailW() - px(32): btnY = y + h - px(16) - px(30)
    Card "WfList", "Workflows", PAGE_PAD, y, lw, h, , 0, 2
    Set c = Btn("btnNewWf", "New workflow", PAGE_PAD + lw - px(16) - px(120), y + px(10), px(120), px(26), "outline")
    c.ControlTipText = "Add a workflow, off until its chain is complete (Ctrl+N)"
    GridHeaders "hdrWf", WorkflowCols(), PAGE_PAD + px(16), y + px(44), w:=lw - px(32)
    Set c = SubF("sfWf", "subGridWorkflows", PAGE_PAD + px(16), y + px(66), lw - px(32), h - px(82)): Anchor c, 0, 2
    Card "WfEdit", "Workflow", x, y, WfDetailW(), h, , 2, 2
    Lbl "lblWfState", "", fx + px(80), y + px(14), fw - px(80) - px(112), px(20), 9, False, CLR_MUTED_FG
    Set c = Btn("btnWfOnOff", "Turn on", fx + fw - px(104), y + px(10), px(104), px(26), "outline")
    hw = (fw - px(8)) \ 2
    FieldLbl "capWfName", "NAME", fx, y + px(44), hw
    Txt "txtWfName", "", fx, y + px(62), hw, px(28)
    FieldLbl "capWfDesc", "DESCRIPTION", fx + hw + px(8), y + px(44), hw
    Txt "txtWfDesc", "", fx + hw + px(8), y + px(62), hw, px(28)
    FieldLbl "capWfTypes", "MESSAGE TYPES (BLANK: ANY; SEPARATE WITH ;)", fx, y + px(98), hw
    Set c = Txt("txtWfTypes", "", fx, y + px(116), hw, px(28))
    c.ControlTipText = "Message types as the type list names them, such as SAFETY MSG; GENADMIN"
    FieldLbl "capWfOrigs", "ORIGINATORS (BLANK: ANY)", fx + hw + px(8), y + px(98), hw
    Set c = Txt("txtWfOrigs", "", fx + hw + px(8), y + px(116), hw, px(28))
    c.ControlTipText = "FM lines, such as COMSUBLANT NORFOLK VA"
    FieldLbl "capWfChain", "CHAIN (THE LAST STEP RELEASES)", fx, y + px(154), fw
    GridHeaders "hdrWfSteps", WorkflowStepCols(), fx, y + px(172), w:=fw
    Set c = SubF("sfWfSteps", "subGridWorkflowSteps", fx, y + px(194), fw, px(150))
    c.TabStop = False
    sx = fx
    Set c = Btn("btnAddStep", "Add step...", sx, y + px(352), px(104), px(28), "outline"): sx = sx + px(112)
    Set c = Btn("btnEditStep", "Change...", sx, y + px(352), px(96), px(28), "outline"): sx = sx + px(104)
    c.ControlTipText = "Rename the picked step or name who holds it (or double-click it)"
    Set c = Btn("btnStepUp", "Up", sx, y + px(352), px(64), px(28), "outline"): sx = sx + px(72)
    Set c = Btn("btnStepDown", "Down", sx, y + px(352), px(64), px(28), "outline"): sx = sx + px(72)
    Set c = Btn("btnRemoveStep", "Remove", sx, y + px(352), px(88), px(28), "outline")
    Lbl "lblWfNote", "", fx, y + px(388), fw, px(48), 9, False, CLR_MUTED_FG
    Set c = Btn("btnWfSave", "Save", fx, btnY, px(88), px(30))
    c.ControlTipText = "Save the name, description and what it matches (Ctrl+S)"
    Btn "btnWfUndo", "Undo", fx + px(96), btnY, px(88), px(30), "outline"
    GroupSince before, "grpWf"
    EmptyState "emptyWf", PAGE_PAD + px(16), y + px(90), lw - px(32), "E8AB", "No workflows yet", _
        "New workflow adds one. A draft takes the most specific active workflow for its type and originator."

    ' -- Release authority --
    Set before = NamesNow()
    lw = RA_LIST_W
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = RaFormW() - px(32)
    Card "RaList", "Release authorities", PAGE_PAD, y, lw, h, , 0, 2
    GridHeaders "hdrRa", AuthorityCols(), PAGE_PAD + px(16), y + px(44), w:=lw - px(32)
    Set c = SubF("sfRa", "subGridAuthorities", PAGE_PAD + px(16), y + px(66), lw - px(32), h - px(66) - px(62)): Anchor c, 0, 2
    Set c = Btn("btnRemoveRa", "Remove designation", PAGE_PAD + px(16), btnY, px(160), px(30), "outline")
    c.ControlTipText = "Take the picked designation away"
    Lbl "lblRaNote", "Releasing needs a designation here, the Designated Release Authority role (Admin: Users) " & _
        "and holding the draft's release step.", PAGE_PAD + px(184), btnY - px(4), lw - px(200), px(36), 8, False, CLR_MUTED_FG
    Card "RaAdd", "Designate", x, y, RaFormW(), px(330), , 2, 0
    FieldLbl "capRaPerson", "PERSON (FROM PERSONNEL)", fx, y + px(44), fw
    Cbo "cboRaPerson", "", "", fx, y + px(62), fw, px(28), 2, "0;" & fw
    FieldLbl "capRaDesig", "RELEASES AS (BLANK: THEIR BILLET)", fx, y + px(98), fw
    Set c = Txt("txtRaDesig", "", fx, y + px(116), fw, px(28))
    c.ControlTipText = "Commander, Chief of Staff, N3..."
    FieldLbl "capRaScope", "MESSAGE TYPES", fx, y + px(152), fw
    Cbo "cboRaScope", "", "", fx, y + px(170), fw, px(28), 2, "0;" & fw
    Chk "chkRaCO", "", fx, y + px(212)
    Lbl "lblRaCO", "The Commander (commanding officer)", fx + px(24), y + px(208), fw - px(24), px(24), 9
    Set c = Btn("btnAddRa", "Designate", fx, y + px(250), px(112), px(30))
    c.ControlTipText = "Make the person a release authority for the message types picked"
    GroupSince before, "grpRa"

    ' -- Command settings --
    Set before = NamesNow()
    fw = px(560) - px(32): fx = PAGE_PAD + px(16)
    Card "CsCard", "Command transmission settings", PAGE_PAD, y, px(560), px(400), , 0, 0
    Lbl "lblCsIntro", "One set of settings governs every message the command releases.", fx, y + px(40), fw, px(20), 9, False, CLR_MUTED_FG
    FieldLbl "capCsCase", "TEXT CASE", fx, y + px(72), px(260)
    Cbo "cboCsCase", "", "upper;""Every message in capitals"";mixed;""As typed""", fx, y + px(90), px(260), px(28), 2, "0;" & px(260)
    FieldLbl "capCsOrig", "ORIGINATOR (FM LINE)", fx, y + px(126), fw
    Set c = Txt("txtCsOrig", "", fx, y + px(144), fw, px(28))
    c.ControlTipText = "Up to 55 characters (NTP 3 603), such as COMSUBLANT NORFOLK VA"
    FieldLbl "capCsMsgId", "MSGID ORIGINATOR", fx, y + px(180), px(260)
    Set c = Txt("txtCsMsgId", "", fx, y + px(198), px(260), px(28))
    c.ControlTipText = "1 to 20 characters (NTP 3 Annex A), such as COMSUBLANT"
    FieldLbl "capCsNext", "NEXT SEQUENCE NUMBER", fx + px(268), y + px(180), fw - px(268)
    Txt "txtCsNext", "", fx + px(268), y + px(198), px(120), px(28)
    Lbl "lblCsSeq", "", fx, y + px(236), fw, px(52), 9, False, CLR_MUTED_FG
    Set c = Btn("btnCsSave", "Save", fx, y + px(354), px(88), px(30))
    c.ControlTipText = "Save the settings (Ctrl+S)"
    Btn "btnCsUndo", "Undo", fx + px(96), y + px(354), px(88), px(30), "outline"
    GroupSince before, "grpCs"
    SaveForm "pgRouting"
End Sub

' ---------------------------------------------------------------------------
' Dialogs
' ---------------------------------------------------------------------------

' New draft: blank, or from an active template (with its description).
Private Sub Build_dlgNewDraft()
    Dim c As Control, w As Long, iw As Long
    w = px(480): iw = w - px(40)
    NewDialog "New draft", w, px(268)
    Ev F, "OnLoad"
    Lbl "lblHead", "New draft", px(20), px(20), iw, px(24), 11, True
    Lbl "lblSub", "Start from a blank draft, or from a template: you fill in its fields next.", px(20), px(46), iw, px(20), 9, False, CLR_MUTED_FG
    FieldLbl "capTemplate", "START FROM", px(20), px(80), iw
    Set c = Cbo("cboTemplate", "", "", px(20), px(98), iw, px(30), 2, "0;" & iw)
    Ev c, "AfterUpdate"
    Lbl "lblTplDesc", "", px(20), px(134), iw, px(48), 9, False, CLR_MUTED_FG
    Set c = Btn("btnStart", "Start", w - px(20) - px(88) - px(8) - px(88), px(218), px(88), px(30))
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(218), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgNewDraft"
End Sub

' A template field's rules (Message Generator's dlgFieldRule): required,
' max length, default, hint for drafters.
Private Sub Build_dlgTemplateField()
    Dim c As Control, x As Long, w As Long, btnY As Long
    x = px(24): w = px(432): btnY = px(332)
    NewDialog "Field rules", px(480), btnY + px(32) + px(24)
    Ev F, "OnLoad"
    Lbl "lblHead", "Field rules", x, px(14), w, px(26), 13, True
    Lbl "lblIntro", "", x, px(44), w, px(50), 9, False, CLR_MUTED_FG
    FieldLbl "capRequired", "REQUIRED", x, px(102), w
    ' The captions hold commas, so each item is quoted: a value list splits on
    ' commas as well as semicolons.
    Cbo "cboRequired", "", "1;""Yes, a draft must fill it"";0;""No, it can stay empty""", x, px(120), px(260), px(28), 2, "0;" & px(260)
    FieldLbl "capMax", "MAX LENGTH", x, px(158), w
    Txt "txtMax", "", x, px(176), px(80), px(28), 10
    Set c = Lbl("lblMaxHint", "characters; 0 for no limit", x + px(92), px(176), w - px(92), px(28), 9, False, CLR_MUTED_FG)
    VCenter c, px(28), 9, False
    FieldLbl "capDefault", "DEFAULT", x, px(214), w
    Set c = Txt("txtDefault", "", x, px(232), w, px(28), 10)
    c.ControlTipText = "Filled in for new drafts; the drafter can change it"
    FieldLbl "capHint", "HINT FOR DRAFTERS", x, px(270), w
    Set c = Txt("txtHint", "", x, px(288), w, px(28), 10)
    c.ControlTipText = "Shown when the drafter fills the field in, up to " & TPL_HINT_MAX & " characters"
    Lbl "lblMsg", "", x, btnY - px(6), w - px(212), px(44), 9, False, CLR_ERROR
    Set c = Btn("btnCancel", "Cancel", x + w - px(200), btnY, px(96), px(32), "outline"): c.Cancel = True
    Btn "btnSave", "Save", x + w - px(96), btnY, px(96), px(32)
    SaveForm "dlgTemplateField"
End Sub

' A workflow step: its name, whether it reviews or releases, who holds it.
Private Sub Build_dlgRouteStep()
    Dim c As Control, w As Long, iw As Long
    w = px(460): iw = w - px(40)
    NewDialog "Step", w, px(330)
    Ev F, "OnLoad"
    Lbl "lblHead", "", px(20), px(20), iw, px(24), 11, True
    Lbl "lblSub", "", px(20), px(46), iw, px(20), 9, False, CLR_MUTED_FG
    FieldLbl "capTitle", "STEP NAME", px(20), px(80), iw
    Txt "txtTitle", "", px(20), px(98), iw, px(28)
    FieldLbl "capAction", "ACTION", px(20), px(134), px(200)
    Cbo "cboAction", "", "review;Review;release;Release", px(20), px(152), px(200), px(28), 2, "0;" & px(200)
    FieldLbl "capPerson", "HELD BY (FROM PERSONNEL)", px(20), px(188), iw
    Cbo "cboPerson", "", "", px(20), px(206), iw, px(28), 2, "0;" & iw
    Lbl "lblMsg", "", px(20), px(242), iw, px(32), 9, False, CLR_ERROR
    Set c = Btn("btnSave", "Save", w - px(20) - px(88) - px(8) - px(88), px(280), px(88), px(30))
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(280), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgRouteStep"
End Sub
