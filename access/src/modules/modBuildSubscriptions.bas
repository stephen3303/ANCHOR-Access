Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Subscriptions page (/subscriptions) and the request dialog.
'
' The page: a note on how requests and mandatory subscriptions work, the
' search box and category filter over the catalogue with the signed-in
' person's status of each, and the picked one on the right: what it is,
' where the person stands (an alert), what it matches and how many
' messages match now, and Request, Pause or Resume, and Withdraw request.
' dlgRequestSub asks how often, by which channels, and why.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, FieldLbl); the kit's modUIKit.

Private Const SUB_DETAIL_W As Long = 5700       ' the detail card, 380 px

Private Function SubListW() As Long
    SubListW = ContentW() - SUB_DETAIL_W - px(12)
End Function

' The tables the page shows, before the page.
Public Sub BuildSubscriptionsGrids()
    BuildGrid "subGridSubs", SubsCols(), SubListW() - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function SubsCols() As Variant
    SubsCols = Array("SubName|Subscription|180|bs", "Category|Category|120|m", "assign_txt|Assigned|132|m", "my_status|Your status|136|p")
End Function

Public Sub BuildSubscriptionsPage()
    BuildSubscriptionsGrids
    Build_pgSubscriptions
End Sub

' The Subscriptions page (see the module's header).
Public Sub Build_pgSubscriptions()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, lw As Long, btnY As Long, k As String
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Subscriptions", "The traffic you follow. Mandatory subscriptions are assigned by your command and can't be paused. " & _
        "A new subscription is a request that your unit's Subscription Approver approves.", PAGE_PAD
    Alert "alInfo", PAGE_PAD, px(84), ContentW(), px(40), "info", "", _
        "New subscriptions go to your unit's Subscription Approver. Mandatory ones are assigned by command and can't be paused."
    ' The banner stretches with the window: its panel and message (the bar and icon stay at its left).
    Anchor F.Controls("alInfo"), 2, 0
    Anchor F.Controls("alInfo_title"), 2, 0
    Anchor F.Controls("alInfo_msg"), 2, 0
    SearchBox "txtSearch", PAGE_PAD, px(136), px(260), "SearchFor", "Search subscriptions...", px(28)
    Set c = Cbo("cboCategory", "", """(All categories)""", PAGE_PAD + px(272), px(136), px(240), px(28))
    c.ControlTipText = "Show one category of the catalogue"
    Ev c, "AfterUpdate"

    y = px(176): h = PAGE_H - y - px(16): lw = SubListW()
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = SUB_DETAIL_W - px(32): btnY = y + h - px(16) - px(30)
    Card "List", "Catalogue", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrList", SubsCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfList"
    Set c = SubF("sfList", "subGridSubs", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), lw - px(32), "EA8F", "No subscriptions match", _
        "No subscription fits the search and the category.", "Show all"

    Card "Detail", "Subscription", x, y, SUB_DETAIL_W, h, , 1, 2
    ' Everything in the card moves with its right edge, as the card does
    ' (AnchorBuilt below); the buttons also keep to its bottom.
    k = BuiltNames()
    Lbl "lblSubName", "", fx, y + px(40), fw, px(24), 11, True
    Lbl "lblSubMeta", "", fx, y + px(64), fw, px(20), 9, False, CLR_MUTED_FG
    Lbl "lblSubDesc", "", fx, y + px(88), fw, px(36), 9, False, CLR_FG
    Alert "alStatus", fx, y + px(132), fw, px(80)
    FieldLbl "capCriteria", "WHAT IT MATCHES", fx, y + px(224), fw
    RO "txtCriteria", "", fx, y + px(242), fw, px(88), 9
    FieldLbl "capMatches", "MATCHING MESSAGES", fx, y + px(338), fw
    Lbl "lblMatches", "", fx, y + px(356), fw, px(20), 9, False, CLR_FG
    Set c = Btn("btnRequest", "Request...", fx, btnY, px(112), px(30))
    c.ControlTipText = "Ask your unit's Subscription Approver for this subscription"
    Ev c, "OnClick"
    Set c = Btn("btnPause", "Pause", fx + px(120), btnY, px(88), px(30), "outline")
    c.ControlTipText = "Stop receiving it for a while, or start again"
    Ev c, "OnClick"
    Set c = Btn("btnWithdraw", "Withdraw request", fx + px(216), btnY, fw - px(216), px(30), "outline")
    c.ControlTipText = "Take back your request before an approver decides it"
    Ev c, "OnClick"
    Set c = Lbl("lblDetailEmpty", "Pick a subscription in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    AnchorBuilt k, 1, 0
    Anchor F.Controls("btnRequest"), 1, 1
    Anchor F.Controls("btnPause"), 1, 1
    Anchor F.Controls("btnWithdraw"), 1, 1
    SaveForm "pgSubscriptions"
End Sub

' The dialogs the Subscriptions page opens (modBuild.BuildAppDialogs).
Public Sub BuildSubscriptionsDialogs()
    Build_dlgRequestSub
End Sub

' Request a subscription: how often (Real-time, Daily Digest, Weekly), the
' channels (In app, Email, Teams) and the justification (required).
Public Sub Build_dlgRequestSub()
    Dim c As Control, w As Long, iw As Long
    w = px(520): iw = w - px(40)
    NewDialog "Request a subscription", w, px(434)
    Ev F, "OnLoad"
    Lbl "lblHead", "", px(20), px(20), iw, px(24), 11, True
    Lbl "lblSub", "", px(20), px(46), iw, px(32), 9, False, CLR_MUTED_FG
    FieldLbl "capFreq", "HOW OFTEN", px(20), px(88), iw
    Segmented "segFreq", "", px(20), px(106), px(360), "rt|Real-time;daily|Daily Digest;weekly|Weekly", px(30)
    FieldLbl "capChannels", "DELIVER TO", px(20), px(148), iw
    Chk "chkInApp", "", px(20), px(170)
    Lbl "lblInApp", "In app", px(44), px(168), px(96), px(20), 9
    Chk "chkEmail", "", px(160), px(170)
    Lbl "lblEmail", "Email", px(184), px(168), px(96), px(20), 9
    Chk "chkTeams", "", px(300), px(170)
    Lbl "lblTeams", "Teams", px(324), px(168), px(96), px(20), 9
    FieldLbl "capWhy", "JUSTIFICATION (REQUIRED)", px(20), px(204), iw
    Set c = Txt("txtWhy", "", px(20), px(222), iw, px(112))
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    Lbl "lblHint", "Say why you need it. The Subscription Approver reads this before deciding.", px(20), px(338), iw, px(20), 9, _
        False, CLR_MUTED_FG
    Set c = Btn("btnSubmit", "Submit for approval", w - px(20) - px(160) - px(8) - px(88), px(384), px(160), px(30))
    c.ControlTipText = "Send the request to your unit's Subscription Approver"
    Ev c, "OnClick"
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(384), px(88), px(30), "outline")
    c.Cancel = True
    Ev c, "OnClick"
    SaveForm "dlgRequestSub"
End Sub
