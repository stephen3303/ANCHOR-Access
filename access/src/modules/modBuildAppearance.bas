Option Compare Database
Option Explicit

' ANCHOR (Access edition) - builds the dlgAppearance palette picker
' (its code-behind is the kit's kit\forms\dlgAppearance.vba).
' BuildUserInterface calls Build_dlgAppearance with the other dialogs;
' ShowAppearance (the kit's modDarkMode) opens it. It began as the kit's
' example module and is now the app's own.
'
' Owner: the app core (build phase 1). Depends on: modPalette; the kit's
' modUIKit.

Public Sub Build_dlgAppearance()
    Dim c As Control, i As Long, k As Long, y As Long
    NewDialog "Appearance", px(360), px(96) + PALETTE_COUNT * px(40)
    Ev F, "OnLoad"
    Lbl "lblHint", "Choose the colors for the whole app.", px(16), px(12), px(328), px(18), 9, False, CLR_MUTED_FG
    For i = 1 To PALETTE_COUNT
        y = px(40) + (i - 1) * px(40)
        Set c = Box("boxRow" & i, px(16), y, px(328), px(34), CLR_SURFACE, CLR_BORDER): c.OnClick = "=Pick(" & i & ")"
        For k = 1 To 3
            Set c = Box("sw" & i & "_" & k, px(28) + (k - 1) * px(22), y + px(8), px(18), px(18), CLR_BORDER, CLR_BORDER)
            c.OnClick = "=Pick(" & i & ")"
        Next
        Link "lblName" & i, PaletteName(i), px(104), y + px(8), px(190), px(18), "=Pick(" & i & ")", 10, CLR_FG
        Set c = Link("lblChk" & i, ChrW(&HE73E), px(306), y + px(8), px(24), px(18), "=Pick(" & i & ")", 10, CLR_PRIMARY)
        c.FontName = "Segoe MDL2 Assets": c.TextAlign = 2
    Next
    Set c = Btn("btnCancel", "Cancel", px(254), px(48) + PALETTE_COUNT * px(40), px(90), px(30), "outline"): c.Cancel = True
    SaveForm "dlgAppearance"
End Sub
