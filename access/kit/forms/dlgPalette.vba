Option Compare Database
Option Explicit

' Command palette for KitPalette (BuildPaletteDialog builds it; modKitPalette
' does the work). It closes when it loses the focus (Deactivate, and a check
' on its Timer), so a click outside it closes it.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    PaletteOpen Me
    Me.txtSearch.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgPalette.Form_Load"
End Sub

Private Sub Form_Deactivate()
    PaletteLostFocus Me
End Sub

Private Sub Form_Timer()
    PaletteTick Me
End Sub

' The search box's handler (SearchBox): narrow the list as the user types.
Public Sub PaletteFilter(ByVal text As String)
    On Error GoTo Err_Handler
    PaletteShow Me, text
    Exit Sub
Err_Handler:
    HandleError "dlgPalette.PaletteFilter"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    PaletteKeyDown Me, KeyCode
    Exit Sub
Err_Handler:
    HandleError "dlgPalette.txtSearch_KeyDown"
End Sub
