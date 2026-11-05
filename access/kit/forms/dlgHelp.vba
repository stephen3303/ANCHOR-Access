Option Compare Database
Option Explicit

' Help window for KitHelp (BuildHelpDialog builds it; modKitHelp does the
' work). OpenArgs: the topic key to show first. Typing in the search box
' narrows the topic list as you type; Enter shows the first match. Esc or
' Close closes it.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    HelpOpen Me, Nz(Me.OpenArgs, "")
    Exit Sub
Err_Handler:
    HandleError "dlgHelp.Form_Load"
End Sub

' The topic shown, so F1 pressed in the help window keeps it (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = HelpCurrentTopic()
End Property

Private Sub txtSearch_Change()
    On Error GoTo Err_Handler
    HelpSearch Me, Nz(Me.txtSearch.Text, "")
    Exit Sub
Err_Handler:
    HandleError "dlgHelp.txtSearch_Change"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        HelpShowFirstMatch Me
    End If
    Exit Sub
Err_Handler:
    HandleError "dlgHelp.txtSearch_KeyDown"
End Sub

Private Sub lstTopics_AfterUpdate()
    On Error GoTo Err_Handler
    If Not IsNull(Me.lstTopics) Then HelpShow Me, Me.lstTopics
    Exit Sub
Err_Handler:
    HandleError "dlgHelp.lstTopics_AfterUpdate"
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgHelp.btnClose_Click"
End Sub
