Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Search page (/search): basic, advanced and saved searches.
'
' A segmented switch picks the mode:
'   Basic     one box of words over the subject, reference, originator and
'             body text, and the results under it (the tall Results card)
'   Advanced  the same words box, a Filters card (reference, originator,
'             type, class, precedence, classification, lifecycle, received
'             from and to, acknowledgement) and the results under it (a
'             shorter Results card)
'   Saved     the person's saved searches, to run or delete
' The two Results cards host the same table (subGridSearch); the page shows
' one at a time. Each mode's controls are listed in a hidden group label
' (grpWords, grpBasic, grpAdv, grpSaved) for the page to show and hide.
' The web app's natural-language search is the plain keyword search here.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, FieldLbl, NamesNow, GroupSince); the kit's modUIKit.

Private Const FILTERS_H As Long = 2040           ' the Filters card, 136 px

' The tables the page shows, before the page.
Public Sub BuildSearchGrids()
    BuildGrid "subGridSearch", SearchCols(), ContentW() - px(32)
    BuildGrid "subGridSavedSearches", SavedCols(), ContentW() - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function SearchCols() As Variant
    ' Subject is last, so it takes the rest of the width and stretches with the window.
    SearchCols = Array("prec|Precedence|84|p|prec_rank", "ref|Reference|140|bs", "originator|Originator|112|s", "msg_type|Type|112|m", _
        "classif_txt|Classification|136|m", "lifecycle|Lifecycle|88|p", "received_txt|Received|100|m|received", "subject|Subject|188|s")
End Function

Private Function SavedCols() As Variant
    SavedCols = Array("SearchName|Name|240|b", "created_txt|Saved|112|m|CreatedOn", "summary|Search|400|m")
End Function

Public Sub BuildSearchPage()
    BuildSearchGrids
    Build_pgSearch
End Sub

' The Search page: the mode switch and Save search over the words box, then
' the mode's cards (see the module's header).
Public Sub Build_pgSearch()
    Dim c As Control, before As Object, x As Long, y As Long, iw As Long, cw As Long, i As Long, fx As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Search", "Find messages by the words in them, or narrow them down by reference, originator, type, class, " & _
        "precedence, classification, lifecycle, dates and acknowledgement. Save a search to run it again.", PAGE_PAD
    Segmented "segMode", "", PAGE_PAD, px(84), px(312), "basic|Basic;advanced|Advanced;saved|Saved", px(32), "ModeChanged"
    Set c = Btn("btnSave", "Save search", PAGE_W - PAGE_PAD - px(120), px(84), px(120), px(32), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Keep these words and filters under a name, to run them again from Saved"
    Ev c, "OnClick"

    Set before = NamesNow()
    Set c = Txt("txtWords", "", PAGE_PAD, px(128), ContentW() - px(112), px(32)): Anchor c, 2, 0
    c.ControlTipText = "Every word must match somewhere: subject, reference, originator or text. Put a phrase in double quotes."
    Ev c, "OnKeyDown"
    Set c = Btn("btnSearch", "Search", PAGE_W - PAGE_PAD - px(104), px(128), px(104), px(32)): Anchor c, 1, 0
    Ev c, "OnClick"
    GroupSince before, "grpWords"

    ' Basic: the tall Results card.
    Set before = NamesNow()
    ResultsCard "Basic", px(172), PAGE_H - px(172) - px(16)
    GroupSince before, "grpBasic"

    ' Advanced: the Filters card and the shorter Results card.
    Set before = NamesNow()
    y = px(172): iw = ContentW() - px(32): fx = PAGE_PAD + px(16)
    cw = (iw - 4 * px(12)) \ 5
    Card "Filters", "", PAGE_PAD, y, ContentW(), FILTERS_H, , 2, 0
    Set c = Btn("btnClear", "Clear filters", PAGE_W - PAGE_PAD - px(120) - px(8) - px(120), px(84), px(120), px(32), "outline")
    Anchor c, 1, 0
    c.ControlTipText = "Empty every filter (the words stay)"
    Ev c, "OnClick"
    For i = 0 To 4
        x = fx + i * (cw + px(12))
        Select Case i
            Case 0
                FieldLbl "capRef", "REFERENCE CONTAINS", x, y + px(16), cw
                Set c = Txt("txtRef", "", x, y + px(34), cw, px(30)): Ev c, "OnKeyDown"
            Case 1
                FieldLbl "capOrig", "ORIGINATOR", x, y + px(16), cw
                Cbo "cboOrig", "", "", x, y + px(34), cw, px(30)
            Case 2
                FieldLbl "capType", "MESSAGE TYPE", x, y + px(16), cw
                Cbo "cboType", "", "", x, y + px(34), cw, px(30)
            Case 3
                FieldLbl "capClass", "CLASS", x, y + px(16), cw
                Cbo "cboClass", "", "", x, y + px(34), cw, px(30), 2, "0;" & cw
            Case 4
                FieldLbl "capPrec", "PRECEDENCE", x, y + px(16), cw
                Cbo "cboPrec", "", "FLASH;IMMEDIATE;PRIORITY;ROUTINE", x, y + px(34), cw, px(30)
        End Select
    Next
    For i = 0 To 4
        x = fx + i * (cw + px(12))
        Select Case i
            Case 0
                FieldLbl "capClassif", "CLASSIFICATION", x, y + px(72), cw
                Cbo "cboClassif", "", "UNCLASSIFIED;CUI;CONFIDENTIAL;SECRET", x, y + px(90), cw, px(30)
            Case 1
                FieldLbl "capLife", "LIFECYCLE", x, y + px(72), cw
                Cbo "cboLife", "", "ACTIVE;AMENDED;CORRECTED;SUPERSEDED;CANCELLED;EXPIRED", x, y + px(90), cw, px(30)
            Case 2
                FieldLbl "capFrom", "RECEIVED FROM", x, y + px(72), cw
                DateBox "dtFrom", "", x, y + px(90), px(30)
            Case 3
                FieldLbl "capTo", "RECEIVED TO", x, y + px(72), cw
                DateBox "dtTo", "", x, y + px(90), px(30)
            Case 4
                FieldLbl "capAck", "ACKNOWLEDGEMENT", x, y + px(72), cw
                Cbo "cboAck", "", "Required;Not required", x, y + px(90), cw, px(30)
        End Select
    Next
    y = y + FILTERS_H + px(12)
    ResultsCard "Adv", y, PAGE_H - y - px(16)
    GroupSince before, "grpAdv"

    ' Saved: the person's saved searches.
    Set before = NamesNow()
    y = px(128)
    Card "Saved", "Saved searches", PAGE_PAD, y, ContentW(), PAGE_H - y - px(16), , 2, 2
    Set c = Btn("btnRun", "Run", PAGE_W - PAGE_PAD - px(16) - px(88) - px(8) - px(88), y + px(8), px(88), px(28)): Anchor c, 1, 0
    c.ControlTipText = "Run the picked search (or double-click it)"
    Ev c, "OnClick"
    Set c = Btn("btnDelete", "Delete", PAGE_W - PAGE_PAD - px(16) - px(88), y + px(8), px(88), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Delete the picked saved search"
    Ev c, "OnClick"
    GridHeaders "hdrSaved", SavedCols(), PAGE_PAD + px(16), y + px(40), w:=iw, sortSf:="sfSaved"
    Set c = SubF("sfSaved", "subGridSavedSearches", PAGE_PAD + px(16), y + px(62), iw, PAGE_H - y - px(16) - px(78)): Anchor c, 2, 2
    GroupSince before, "grpSaved"
    EmptyState "emptySaved", PAGE_PAD + px(16), y + px(90), iw, "E728", "No saved searches", _
        "Run a search, then click Save search to keep it here under a name."
    SaveForm "pgSearch"
End Sub

' A Results card nm at y, h tall: Open in its header, the table's captions
' and the table (sf<nm>), and its empty state (empty<nm>).
Private Sub ResultsCard(ByVal nm As String, ByVal y As Long, ByVal h As Long)
    Dim c As Control, iw As Long
    iw = ContentW() - px(32)
    Card nm, "Results", PAGE_PAD, y, ContentW(), h, , 2, 2
    Set c = Btn("btnOpen" & nm, "Open", PAGE_W - PAGE_PAD - px(16) - px(88), y + px(8), px(88), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Read the picked message in full (or double-click it)"
    Ev c, "OnClick"
    GridHeaders "hdr" & nm, SearchCols(), PAGE_PAD + px(16), y + px(40), w:=iw, sortSf:="sf" & nm
    Set c = SubF("sf" & nm, "subGridSearch", PAGE_PAD + px(16), y + px(62), iw, h - px(78)): Anchor c, 2, 2
    EmptyState "empty" & nm, PAGE_PAD + px(16), y + px(90), iw, "E721", "Search ANCHOR's messages", _
        "Type words and click Search. Every word must match somewhere in a message."
End Sub
