Option Compare Database
Option Explicit

' Popover for KitPopover (BuildPopoverDialog builds it; modMenu does the
' work). It closes when it loses the focus, so a click outside it closes it:
' Deactivate, and its timer checking it is still the active form (MenuTick).

Private Sub Form_Load()
    On Error GoTo Err_Handler
    PopoverShow Me
    Me.txtKeys.SetFocus
    MenuStarted Me
    Exit Sub
Err_Handler:
    HandleError "dlgPopover.Form_Load"
End Sub

Private Sub Form_Deactivate()
    MenuLostFocus Me
End Sub

Private Sub Form_Close()
    PopoverClosed
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    MenuTick Me
End Sub

Private Sub txtKeys_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    PopoverKey Me, KeyCode
    Exit Sub
Err_Handler:
    HandleError "dlgPopover.txtKeys_KeyDown"
End Sub
