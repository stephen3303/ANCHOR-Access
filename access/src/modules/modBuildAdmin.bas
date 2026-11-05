Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Admin page (/admin): subscription approvals,
' mandatory alert rules, command access, users and their roles, and the
' message types. Its code-behind is forms\pgAdmin.vba; its data modAdmin and
' modSubscriptions (PendingSubsSql, ApproveSubscription, DenySubscription).
'
' Three tabs, a Segmented control under the header (segTab):
'   Approvals & Access  the requests waiting for a decision (Approve, Deny),
'                       the mandatory alert rules and the command access
'                       groups, both read only (grpA)
'   Users               add a person from the Personnel file; the people with
'                       ANCHOR access on the left, the picked person's profile,
'                       roles and access on the right with Save (grpU)
'   Message types       the types by category with the picked one on the
'                       right (Add, Save, Retire or Restore), and the
'                       communications classes, read only (grpT)
' Each tab's controls are listed in a hidden label (GroupSince) that the page
' shows and hides together. Users and Message types are for System Admins;
' a Subscription Approver sees only the first tab, with no tabs shown.
' The web app's Routing & Release tab is the Routing page (btnRouting).
'
' The list pages' pattern: the list card on the left and a 360px detail card
' on the right.
'
' Owner: the Admin page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, NamesNow, GroupSince, FieldLbl); the kit's
' modPersonnelUI, modUIKit.

Private Const DETAIL_W As Long = 5400            ' the detail cards, 360 px
Private Const RULES_W As Long = 8100             ' the alert rules card, 540 px
Private Const TAB_Y As Long = 1860               ' the tabs' content top, 124 px
Private Const ROLE_SLOTS As Long = 8             ' role check boxes on the Users tab (tblRoles has 7)

' The list cards' width beside a detail card.
Private Function AdmListW() As Long
    AdmListW = ContentW() - DETAIL_W - px(12)
End Function

' The detail cards' left edge.
Private Function AdmDetailX() As Long
    AdmDetailX = PAGE_PAD + AdmListW() + px(12)
End Function

' The access groups card's width, beside the rules card.
Private Function AccessW() As Long
    AccessW = ContentW() - RULES_W - px(12)
End Function

' modBuild.BuildPages: the tables first, then the page.
Public Sub BuildAdminPage()
    BuildAdminGrids
    Build_pgAdmin
End Sub

' Every table the Admin page shows, before the page that hosts them.
Public Sub BuildAdminGrids()
    BuildGrid "subGridAdmQueue", QueueCols(), ContentW() - px(32)
    BuildGrid "subGridAdmRules", RuleCols(), RULES_W - px(32), 2
    BuildGrid "subGridAdmAccess", AccessCols(), AccessW() - px(32), 2
    BuildGrid "subGridAdmUsers", UserCols(), AdmListW() - px(32)
    BuildGrid "subGridAdmTypes", TypeCols(), AdmListW() - px(32)
    BuildGrid "subGridAdmClasses", ClassCols(), AdmListW() - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function QueueCols() As Variant
    QueueCols = Array("PersonName|Requester|150|b", "SubName|Subscription|220|", "Frequency|Frequency|112|p", _
        "requested_txt|Requested|112|m|RequestedOn", "Justification|Justification|200|")
End Function

Private Function RuleCols() As Variant
    RuleCols = Array("RuleKind|Rule|112|bw", "MatchText|Matches|120|w", "Audience|Audience|112|w", "ack_txt|Ack|72|m", _
        "Escalate|Escalates to|80|")
End Function

Private Function AccessCols() As Variant
    AccessCols = Array("GroupName|Group|112|bw", "Basis|Basis|104|mw", "Members|Members|72|r", "Scope|Scope|100|")
End Function

Private Function UserCols() As Variant
    UserCols = Array("person|Name|180|b", "billet|Billet|150|m", "roles|Roles|170|", "state|Status|88|p")
End Function

Private Function TypeCols() As Variant
    TypeCols = Array("Category|Category|160|m", "TypeCode|Code|140|b", "TypeName|Name|200|", "state|Status|88|p")
End Function

Private Function ClassCols() As Variant
    ClassCols = Array("ClassCode|Class|56|b", "ClassName|Name|200|", "Description|Description|200|m")
End Function

' The page: the header, the tabs, then each tab's controls (see the top).
Public Sub Build_pgAdmin()
    Dim before As Object, c As Control
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Admin", "Approve subscription requests, review the rules that make delivery mandatory, and manage ANCHOR users, " & _
        "their roles and the message types.", PAGE_PAD
    Segmented "segTab", "", PAGE_PAD, px(80), px(456), "access|Approvals & Access;users|Users;types|Message types", px(32), "TabChanged"
    ' Admin: Routing (the drafting chains, release authority, command
    ' settings) is a page of its own; System Admins open it from here.
    Set c = Btn("btnRouting", "Routing", PAGE_W - PAGE_PAD - px(112), px(24), px(112), px(32), "outline")
    c.ControlTipText = "Workflows, release authority and command settings for drafting": Anchor c, 1, 0

    Set before = NamesNow()
    BuildAccessTab
    GroupSince before, "grpA"
    Set before = NamesNow()
    BuildUsersTab
    GroupSince before, "grpU"
    Set before = NamesNow()
    BuildTypesTab
    GroupSince before, "grpT"
    ' The empty states are built hidden, so no group lists them: the page
    ' shows each with its tab.
    BuildEmptyStates
    SaveForm "pgAdmin"
End Sub

' Approvals & Access: the queue across the top, the rules and the access
' groups side by side under it.
Private Sub BuildAccessTab()
    Dim c As Control, x As Long, y As Long, w As Long, h As Long
    y = TAB_Y: w = ContentW(): h = px(212)
    Card "Queue", "Subscription approval queue", PAGE_PAD, y, w, h, , 2, 0
    Set c = Btn("btnDeny", "Deny", PAGE_PAD + w - px(16) - px(88), y + px(8), px(88), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Deny the picked request, with a reason the requester sees"
    Set c = Btn("btnApprove", "Approve", PAGE_PAD + w - px(16) - px(88) - px(8) - px(96), y + px(8), px(96), px(28)): Anchor c, 1, 0
    c.ControlTipText = "Approve the picked request: its traffic starts reaching the requester"
    Lbl "lblQueueNote", "Requests aren't active until approved. Pick a request, then Approve or Deny.", _
        PAGE_PAD + px(16), y + px(40), w - px(32), px(18), 9, False, CLR_MUTED_FG
    GridHeaders "hdrQueue", QueueCols(), PAGE_PAD + px(16), y + px(66), w:=w - px(32), sortSf:="sfQueue"
    Set c = SubF("sfQueue", "subGridAdmQueue", PAGE_PAD + px(16), y + px(88), w - px(32), h - px(104)): Anchor c, 2, 0

    y = y + h + px(12): h = PAGE_H - y - px(16)
    Card "Rules", "Mandatory alert rules", PAGE_PAD, y, RULES_W, h, , 0, 2
    Lbl "lblRulesNote", "Read only: set by command policy", PAGE_PAD + RULES_W - px(16) - px(220), y + px(12), px(220), px(18), _
        8, False, CLR_MUTED_FG
    F.Controls("lblRulesNote").TextAlign = 3
    GridHeaders "hdrRules", RuleCols(), PAGE_PAD + px(16), y + px(44), w:=RULES_W - px(32)
    Set c = SubF("sfRules", "subGridAdmRules", PAGE_PAD + px(16), y + px(66), RULES_W - px(32), h - px(82)): Anchor c, 0, 2

    x = PAGE_PAD + RULES_W + px(12)
    Card "Access", "Command access groups", x, y, AccessW(), h, , 2, 2
    Lbl "lblAccessNote", "Changes need a System Admin", x + AccessW() - px(16) - px(200), y + px(12), px(200), px(18), _
        8, False, CLR_MUTED_FG
    F.Controls("lblAccessNote").TextAlign = 3: Anchor F.Controls("lblAccessNote"), 1, 0
    GridHeaders "hdrAccess", AccessCols(), x + px(16), y + px(44), w:=AccessW() - px(32)
    Set c = SubF("sfAccess", "subGridAdmAccess", x + px(16), y + px(66), AccessW() - px(32), h - px(82)): Anchor c, 2, 2
End Sub

' Users: the picker and Add person over the list on the left, the picked
' person on the right.
Private Sub BuildUsersTab()
    Dim c As Control, i As Long, x As Long, y As Long, h As Long, fx As Long, fw As Long, hw As Long, lw As Long, ry As Long
    Dim k As String
    y = TAB_Y: lw = AdmListW()
    Set c = PersonPickBox("txtPerson", "", PAGE_PAD, y, px(300), px(30))
    Ev c, "OnKeyDown"
    Set c = Btn("btnAdd", "Add person", PAGE_PAD + px(308), y, px(112), px(30), "outline")
    c.ControlTipText = "Give the picked person ANCHOR access as Message Viewer and Subscription Requester"
    Lbl "lblAddHint", "Only people in the Personnel file can be picked.", PAGE_PAD + px(432), y, lw - px(432), px(30), _
        9, False, CLR_MUTED_FG
    ' Everything over and in the detail card keeps to the window's right edge
    ' as the list widens (AnchorBuilt); the Save buttons also keep to its bottom.
    x = AdmDetailX()
    k = BuiltNames()
    Toggle "tgShowOff", "", x + px(4), y + px(5), px(36), px(20), "ShowOffChanged"
    Link "lbltgShowOff", "Show deactivated people", x + px(52), y + px(1), DETAIL_W - px(52), px(28), _
        "=ToggleSwitch([Form],""tgShowOff"")", 9, CLR_FG
    AnchorBuilt k, 1, 0

    y = y + px(44): h = PAGE_H - y - px(16)
    Card "Users", "Users", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrUsers", UserCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfUsers"
    Set c = SubF("sfUsers", "subGridAdmUsers", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2

    ' The picked person.
    fx = x + px(16): fw = DETAIL_W - px(32): hw = (fw - px(8)) \ 2
    Card "Person", "Person", x, y, DETAIL_W, h, , 1, 2
    k = BuiltNames()
    Lbl "lblName", "", fx, y + px(44), fw, px(24), 11, True
    Lbl "lblWho", "", fx, y + px(68), fw, px(18), 9, False, CLR_MUTED_FG
    FieldLbl "lblBilletCap", "BILLET", fx, y + px(96), fw
    Txt "txtBillet", "", fx, y + px(114), fw, px(28)
    FieldLbl "lblUnitCap", "UNIT", fx, y + px(150), hw
    FieldLbl "lblCommunityCap", "COMMUNITY", fx + hw + px(8), y + px(150), hw
    Txt "txtUnit", "", fx, y + px(168), hw, px(28)
    Txt "txtCommunity", "", fx + hw + px(8), y + px(168), hw, px(28)
    FieldLbl "lblLocationCap", "LOCATION", fx, y + px(204), hw
    FieldLbl "lblDsnCap", "DSN", fx + hw + px(8), y + px(204), hw
    Txt "txtLocation", "", fx, y + px(222), hw, px(28)
    Txt "txtDsn", "", fx + hw + px(8), y + px(222), hw, px(28)
    FieldLbl "lblRolesCap", "ROLES", fx, y + px(260), fw
    ' Eight slots, two to a row; Form_Load names them from tblRoles.
    For i = 1 To ROLE_SLOTS
        ry = y + px(280) + ((i - 1) \ 2) * px(24)
        Chk "chkRole" & i, "", fx + ((i - 1) Mod 2) * (hw + px(8)), ry + px(2)
        Set c = Lbl("lblRole" & i, "", fx + ((i - 1) Mod 2) * (hw + px(8)) + px(24), ry, hw - px(24), px(20), 9, False, CLR_FG)
        Ev c, "OnClick"
    Next
    Toggle "tgActive", "", fx + px(4), y + px(380), px(36), px(20)
    Link "lbltgActive", "Can use ANCHOR", fx + px(52), y + px(376), fw - px(52), px(28), "=ToggleSwitch([Form],""tgActive"")", 9, CLR_FG
    Set c = Lbl("lblPersonEmpty", "Pick someone in the list to see their profile and roles.", fx, y + px(120), fw, px(20), _
        9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    AnchorBuilt k, 1, 0
    Set c = Btn("btnSave", "Save", fx, y + h - px(46), px(88), px(30)): Anchor c, 1, 1
    c.ControlTipText = "Save this person's profile, roles and access (Ctrl+S)"
    Set c = Btn("btnUndo", "Undo changes", fx + px(96), y + h - px(46), px(120), px(30), "outline"): Anchor c, 1, 1
    c.ControlTipText = "Put back what was last saved"
End Sub

' Message types: the types and, under them, the classes on the left; the
' picked type on the right.
Private Sub BuildTypesTab()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, lw As Long, ch As Long, k As String
    y = TAB_Y: lw = AdmListW(): ch = px(236)
    h = PAGE_H - y - px(16) - ch - px(12)
    Card "Types", "Message types", PAGE_PAD, y, lw, h, , 2, 2
    Set c = Btn("btnAddType", "Add type", PAGE_PAD + lw - px(16) - px(96), y + px(8), px(96), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Add a message type"
    GridHeaders "hdrTypes", TypeCols(), PAGE_PAD + px(16), y + px(44), w:=lw - px(32), sortSf:="sfTypes"
    Set c = SubF("sfTypes", "subGridAdmTypes", PAGE_PAD + px(16), y + px(66), lw - px(32), h - px(82)): Anchor c, 2, 2

    Card "Classes", "Communications classes", PAGE_PAD, y + h + px(12), lw, ch, , 2, 1
    Set c = Lbl("lblClassesNote", "Read only", PAGE_PAD + lw - px(16) - px(160), y + h + px(24), px(160), px(18), 8, False, CLR_MUTED_FG)
    c.TextAlign = 3: Anchor c, 1, 1
    k = BuiltNames()
    GridHeaders "hdrClasses", ClassCols(), PAGE_PAD + px(16), y + h + px(56), w:=lw - px(32)
    ' The headers move down with the card (it keeps to the window's bottom);
    ' their own left/right anchoring stays as GridHeaders set it.
    For Each c In F.Controls
        If InStr(1, k, ";" & c.Name & ";", vbTextCompare) = 0 Then c.VerticalAnchor = 1
    Next
    Set c = SubF("sfClasses", "subGridAdmClasses", PAGE_PAD + px(16), y + h + px(78), lw - px(32), ch - px(82)): Anchor c, 2, 1

    x = AdmDetailX(): fx = x + px(16): fw = DETAIL_W - px(32): h = PAGE_H - y - px(16)
    Card "Type", "Message type", x, y, DETAIL_W, h, , 1, 2
    k = BuiltNames()     ' the type's fields keep to the right edge (AnchorBuilt)
    Lbl "lblTypeHead", "", fx, y + px(44), fw, px(24), 11, True
    Lbl "lblTypeState", "", fx, y + px(68), fw, px(18), 9, False, CLR_MUTED_FG
    FieldLbl "lblTypeCodeCap", "CODE", fx, y + px(96), fw
    Txt "txtTypeCode", "", fx, y + px(114), fw, px(28)
    FieldLbl "lblTypeNameCap", "NAME", fx, y + px(150), fw
    Txt "txtTypeName", "", fx, y + px(168), fw, px(28)
    FieldLbl "lblTypeCatCap", "CATEGORY", fx, y + px(204), fw
    Set c = Cbo("cboTypeCat", "", "", fx, y + px(222), fw, px(28))
    c.LimitToList = False                   ' a new category is typed in
    c.ControlTipText = "Pick a category, or type a new one"
    FieldLbl "lblTypeDescCap", "DESCRIPTION", fx, y + px(258), fw
    Set c = Txt("txtTypeDesc", "", fx, y + px(276), fw, px(88))
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    Lbl "lblTypeHint", "The code is what messages name, so it can't change once added. Retire a type no longer used: " & _
        "messages of that type keep it.", fx, y + px(372), fw, px(48), 8, False, CLR_MUTED_FG
    Set c = Lbl("lblTypeEmpty", "Pick a message type in the list, or Add type.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    AnchorBuilt k, 1, 0
    Set c = Btn("btnSaveType", "Save", fx, y + h - px(46), px(88), px(30)): Anchor c, 1, 1
    c.ControlTipText = "Save the message type (Ctrl+S)"
    Set c = Btn("btnRetire", "Retire", fx + px(96), y + h - px(46), px(96), px(30), "outline"): Anchor c, 1, 1
    c.ControlTipText = "Stop offering this type for new traffic, or offer it again"
End Sub

' The lists' empty states, over each list (built hidden).
Private Sub BuildEmptyStates()
    Dim y As Long, lw As Long, x As Long
    y = TAB_Y: lw = AdmListW()
    EmptyState "emptyQueue", PAGE_PAD + px(16), y + px(60), ContentW() - px(32), "E930", "No requests waiting", _
        "Subscription requests from your unit show here until someone approves or denies them."
    EmptyState "emptyRules", PAGE_PAD + px(16), y + px(212) + px(12) + px(56), RULES_W - px(32), "E7BA", "No alert rules", _
        "The rules that make delivery mandatory show here."
    x = PAGE_PAD + RULES_W + px(12)
    EmptyState "emptyAccess", x + px(16), y + px(212) + px(12) + px(56), AccessW() - px(32), "E716", "No access groups", _
        "The command access groups show here."
    EmptyState "emptyUsers", PAGE_PAD + px(16), y + px(44) + px(90), lw - px(32), "E716", "Nobody here yet", _
        "Pick someone from the Personnel file in the box above, then Add person."
    EmptyState "emptyTypes", PAGE_PAD + px(16), y + px(56), lw - px(32), "E8A5", "No message types", _
        "Add the types messages are sorted by, with Add type."
End Sub
