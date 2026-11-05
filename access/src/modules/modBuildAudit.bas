Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Audit page (/audit): the append-only audit
' trail and its hash chain check, for System Admins. Its code-behind is
' forms\pgAudit.vba; the chain itself is modAudit's.
'
' Four tiles (entries, failed, warnings, the chain's status), the result of
' the last Verify in an alert under them, the filters (search over resource
' and details, action, result, user, a day range), then the trail on the
' left, oldest or newest first, and the picked entry in full on the right,
' every column, the hashes included. Verify chain and Export (CSV) sit at
' the top right.
'
' Owner: the Audit page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW); the kit's modUIKit.

Private Const AUD_DETAIL_W As Long = 5400       ' the entry card, 360 px

Private Function AudListW() As Long
    AudListW = ContentW() - AUD_DETAIL_W - px(12)
End Function

' modBuild.BuildPages: the table first, then the page.
Public Sub BuildAuditPage()
    BuildGrid "subGridAudit", AuditCols(), AudListW() - px(32)
    Build_pgAudit
End Sub

' "field|caption|width px|options[|sort field]"
Private Function AuditCols() As Variant
    AuditCols = Array("AuditID|#|48|m", "when_txt|When|112|m|LoggedAt", "UserName|User|100|", "Action|Action|104|", _
        "Resource|Resource|152|s", "Result|Result|72|p")
End Function

Public Sub Build_pgAudit()
    Dim c As Control, i As Long, x As Long, y As Long, h As Long, tw As Long, lw As Long, fx As Long, fw As Long
    Dim tiles As Variant, labels As Variant, glyphs As Variant
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnTimer"
    PageHeader "Audit", "Every audited action, append-only. Each entry carries the hash of the entry before it, so Verify " & _
        "chain shows whether any entry was changed, reordered or removed.", PAGE_PAD
    Set c = Btn("btnExport", "Export", PAGE_W - PAGE_PAD - px(88), px(24), px(88), px(32), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Save the entries the filters show as a CSV file (the export is audited too)"
    Set c = Btn("btnVerify", "Verify chain", PAGE_W - PAGE_PAD - px(88) - px(8) - px(128), px(24), px(128), px(32)): Anchor c, 1, 0
    c.ControlTipText = "Recompute every entry's hash from the first to the last"

    tiles = Array("tileEntries", "tileFailed", "tileWarnings", "tileChain")
    labels = Array("Entries in the trail", "Failed", "Warnings", "Chain")
    glyphs = Array("E81C", "E783", "E7BA", "E72E")
    tw = (ContentW() - 3 * px(12)) \ 4
    For i = 0 To 3
        x = PAGE_PAD + i * (tw + px(12))
        StatTile CStr(tiles(i)), CStr(labels(i)), "", x, px(84), IIf(i = 3, ContentW() - 3 * (tw + px(12)), tw), px(104), _
            caption:=" ", glyph:=CStr(glyphs(i)), fmt:=IIf(i = 3, "", "#,##0")
    Next
    Alert "alVerify", PAGE_PAD, px(200), ContentW(), px(40)

    ' The filters.
    y = px(252): x = PAGE_PAD
    SearchBox "txtSearch", x, y, px(168), "SearchFor", "Search the trail", px(28)
    x = x + px(176)
    Set c = Cbo("cboAction", "", "", x, y, px(128), px(28), 2, "0;" & px(128)): Ev c, "AfterUpdate"
    c.ControlTipText = "Only this action"
    x = x + px(136)
    Set c = Cbo("cboResult", "", "", x, y, px(104), px(28), 2, "0;" & px(104)): Ev c, "AfterUpdate"
    c.ControlTipText = "Only this result"
    x = x + px(112)
    Set c = Cbo("cboUser", "", "", x, y, px(128), px(28), 2, "0;" & px(128)): Ev c, "AfterUpdate"
    c.ControlTipText = "Only this user"
    x = x + px(140)
    Lbl "lblFrom", "From", x, y, px(36), px(28), 9, False, CLR_MUTED_FG
    Set c = DateBox("dtFrom", "", x + px(40), y, px(28)): Ev c, "AfterUpdate"
    Ev F.Controls("cal_dtFrom"), "OnClick"
    x = x + px(40) + DATEFIELD_W + px(8)
    Lbl "lblTo", "to", x, y, px(20), px(28), 9, False, CLR_MUTED_FG
    Set c = DateBox("dtTo", "", x + px(28), y, px(28)): Ev c, "AfterUpdate"
    Ev F.Controls("cal_dtTo"), "OnClick"

    ' The trail.
    y = px(292): h = PAGE_H - y - px(16): lw = AudListW()
    Card "List", "Audit trail", PAGE_PAD, y, lw, h, , 2, 2
    Segmented "segOrder", "", PAGE_PAD + lw - px(16) - px(224), y + px(6), px(224), "old|Oldest first;new|Newest first", px(28), "OrderChanged"
    For Each c In F.Controls
        If c.Name Like "segOrder*" Then Anchor c, 1, 0
    Next
    GridHeaders "hdrList", AuditCols(), PAGE_PAD + px(16), y + px(44), w:=lw - px(32), sortSf:="sfList"
    Set c = SubF("sfList", "subGridAudit", PAGE_PAD + px(16), y + px(66), lw - px(32), h - px(82)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(66), lw - px(32), "E81C", "No matching entries", _
        "No entry fits the search, the filters and the dates.", "Show all"

    ' The picked entry.
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = AUD_DETAIL_W - px(32)
    Card "Detail", "Entry", x, y, AUD_DETAIL_W, h, , 1, 2
    Set c = RO("txtDetail", "", fx, y + px(40), fw, h - px(40) - px(62), 9)
    c.ScrollBars = 2: Anchor c, 1, 2
    Set c = Btn("btnCopy", "Copy", fx, y + h - px(46), px(80), px(30), "outline"): Anchor c, 1, 1
    c.ControlTipText = "Copy the whole entry, to paste into an email or a report"
    Lbl "lblDetailNote", "Entries can't be changed or removed.", fx + px(88), y + h - px(46), fw - px(88), px(30), 8, False, CLR_MUTED_FG
    Anchor F.Controls("lblDetailNote"), 1, 1
    Set c = Lbl("lblDetailEmpty", "Pick an entry in the list to see all of it here.", fx, y + px(120), fw, px(20), 9, False, _
        CLR_MUTED_FG)
    c.TextAlign = 2
    SaveForm "pgAudit"
End Sub
