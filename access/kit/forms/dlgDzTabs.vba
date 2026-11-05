Option Compare Database
Option Explicit

' Form designer tab order (BuildDesignerTabsDialog, modKitDesigner): the
' elements that take the focus, moved up and down or put back in reading
' order; OK gives them that order. RunDialog's result: "ok" when it did.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    DesignerTabsLoad Me
    Exit Sub
Err_Handler:
    HandleError "dlgDzTabs.Form_Load"
End Sub

Private Sub lstTabs_AfterUpdate()
    On Error GoTo Err_Handler
    DesignerTabsPicked Me
    Exit Sub
Err_Handler:
    HandleError "dlgDzTabs.lstTabs_AfterUpdate"
End Sub

Private Sub btnUp_Click()
    On Error GoTo Err_Handler
    DesignerTabsMove Me, -1
    Exit Sub
Err_Handler:
    HandleError "dlgDzTabs.btnUp_Click"
End Sub

Private Sub btnDown_Click()
    On Error GoTo Err_Handler
    DesignerTabsMove Me, 1
    Exit Sub
Err_Handler:
    HandleError "dlgDzTabs.btnDown_Click"
End Sub

Private Sub btnReading_Click()
    On Error GoTo Err_Handler
    DesignerTabsReading Me
    Exit Sub
Err_Handler:
    HandleError "dlgDzTabs.btnReading_Click"
End Sub

Private Sub btnOK_Click()
    On Error GoTo Err_Handler
    DesignerTabsSave Me
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgDzTabs.btnOK_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgDzTabs.btnCancel_Click"
End Sub
